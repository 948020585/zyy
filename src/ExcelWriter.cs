using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace CertPhotoSorter
{
    /// <summary>
    /// Excel 写入功能，用于更新原始Excel文件中的匹配状态
    /// </summary>
    internal static class ExcelWriter
    {
        /// <summary>
        /// 在原始Excel文件中添加/更新匹配状态列
        /// </summary>
        public static void UpdateMatchStatus(
            string excelPath,
            string worksheet,
            List<PersonMatchStatus> matchStatusList)
        {
            if (string.IsNullOrWhiteSpace(excelPath))
                throw new ArgumentException("Excel路径不能为空", "excelPath");
            if (matchStatusList == null || matchStatusList.Count == 0)
                return;

            var backupPath = CreateBackup(excelPath);

            try
            {
                var extProps = ExcelHelper.GetWriteExtendedProperties(excelPath);

                Exception lastException = null;
                foreach (var provider in ExcelHelper.Providers)
                {
                    try
                    {
                        var connStr = ExcelHelper.BuildConnectionString(provider, excelPath, extProps);
                        using (var conn = new OleDbConnection(connStr))
                        {
                            conn.Open();
                            var sheet = ExcelHelper.ResolveWorksheetName(conn, worksheet, true);
                            EnsureColumnsExist(conn, sheet);
                            UpdateData(conn, sheet, matchStatusList);
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                    }
                }

                throw new InvalidOperationException("无法通过 OLEDB 写入 Excel。最后错误: " + lastException, lastException);
            }
            catch
            {
                RestoreBackup(backupPath, excelPath);
                throw;
            }
        }

        private static string CreateBackup(string excelPath)
        {
            var dir = Path.GetDirectoryName(excelPath);
            var nameWithoutExt = Path.GetFileNameWithoutExtension(excelPath);
            var ext = Path.GetExtension(excelPath);
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            var backupPath = Path.Combine(dir, nameWithoutExt + "_backup_" + timestamp + ext);
            File.Copy(excelPath, backupPath, true);
            return backupPath;
        }

        private static void RestoreBackup(string backupPath, string excelPath)
        {
            if (File.Exists(backupPath))
            {
                File.Delete(excelPath);
                File.Move(backupPath, excelPath);
            }
        }

        private static void EnsureColumnsExist(OleDbConnection conn, string sheet)
        {
            var columns = GetExistingColumns(conn, sheet);

            if (!columns.Contains(Texts.ColMatched))
            {
                TryAddColumn(conn, sheet, Texts.ColMatched, "VARCHAR(10)");
            }

            if (!columns.Contains(Texts.ColMatchReason))
            {
                TryAddColumn(conn, sheet, Texts.ColMatchReason, "VARCHAR(255)");
            }
        }

        private static HashSet<string> GetExistingColumns(OleDbConnection conn, string sheet)
        {
            var columns = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT TOP 1 * FROM [" + sheet + "]";
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        var dt = new DataTable();
                        da.Fill(dt);
                        foreach (DataColumn col in dt.Columns)
                        {
                            columns.Add(col.ColumnName);
                        }
                    }
                }
            }
            catch { }
            return columns;
        }

        private static void TryAddColumn(OleDbConnection conn, string sheet, string columnName, string dataType)
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = string.Format(
                        CultureInfo.InvariantCulture,
                        "ALTER TABLE [{0}] ADD [{1}] {2}",
                        sheet, columnName, dataType);
                    cmd.ExecuteNonQuery();
                }
            }
            catch
            {
                // OLEDB 对 Excel 的 DDL 支持有限，忽略添加列失败
            }
        }

        private static void UpdateData(OleDbConnection conn, string sheet, List<PersonMatchStatus> matchStatusList)
        {
            // 先检查目标列是否存在
            var columns = GetExistingColumns(conn, sheet);
            var hasMatchedCol = columns.Contains(Texts.ColMatched);
            var hasReasonCol = columns.Contains(Texts.ColMatchReason);

            if (!hasMatchedCol && !hasReasonCol) return;

            foreach (var status in matchStatusList)
            {
                if (string.IsNullOrWhiteSpace(status.Id)) continue;

                try
                {
                    using (var cmd = conn.CreateCommand())
                    {
                        var matchedValue = status.Matched ? Texts.MatchValueYes : Texts.MatchValueNo;
                        var reasonValue = status.MatchReason ?? string.Empty;

                        if (hasMatchedCol && hasReasonCol)
                        {
                            cmd.CommandText = string.Format(
                                CultureInfo.InvariantCulture,
                                "UPDATE [{0}] SET [{1}] = ?, [{2}] = ? WHERE [{3}] = ?",
                                sheet, Texts.ColMatched, Texts.ColMatchReason, Texts.ColId);
                            cmd.Parameters.Add(new OleDbParameter("?", matchedValue));
                            cmd.Parameters.Add(new OleDbParameter("?", reasonValue));
                            cmd.Parameters.Add(new OleDbParameter("?", status.Id));
                        }
                        else if (hasMatchedCol)
                        {
                            cmd.CommandText = string.Format(
                                CultureInfo.InvariantCulture,
                                "UPDATE [{0}] SET [{1}] = ? WHERE [{2}] = ?",
                                sheet, Texts.ColMatched, Texts.ColId);
                            cmd.Parameters.Add(new OleDbParameter("?", matchedValue));
                            cmd.Parameters.Add(new OleDbParameter("?", status.Id));
                        }
                        else
                        {
                            cmd.CommandText = string.Format(
                                CultureInfo.InvariantCulture,
                                "UPDATE [{0}] SET [{1}] = ? WHERE [{2}] = ?",
                                sheet, Texts.ColMatchReason, Texts.ColId);
                            cmd.Parameters.Add(new OleDbParameter("?", reasonValue));
                            cmd.Parameters.Add(new OleDbParameter("?", status.Id));
                        }

                        cmd.ExecuteNonQuery();
                    }
                }
                catch
                {
                    // 忽略单行更新失败，继续处理其他行
                }
            }
        }
    }
}
