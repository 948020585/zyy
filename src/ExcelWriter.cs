using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;

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

            // 创建临时备份
            var backupPath = CreateBackup(excelPath);

            try
            {
                var extProps = GetExtendedProperties(excelPath);
                var providers = new[]
                {
                    "Microsoft.ACE.OLEDB.16.0",
                    "Microsoft.ACE.OLEDB.12.0",
                    "Microsoft.Jet.OLEDB.4.0"
                };

                Exception lastException = null;
                for (int i = 0; i < providers.Length; i++)
                {
                    var provider = providers[i];
                    OleDbConnection conn = null;
                    try
                    {
                        var connStr = string.Format(
                            CultureInfo.InvariantCulture,
                            "Provider={0};Data Source={1};Extended Properties='{2}'",
                            provider,
                            excelPath,
                            extProps);

                        conn = new OleDbConnection(connStr);
                        conn.Open();

                        // 确定工作表名称
                        var sheet = ResolveWorksheetName(conn, worksheet);

                        // 添加新列（如果不存在）
                        EnsureColumnsExist(conn, sheet);

                        // 更新数据
                        UpdateData(conn, sheet, matchStatusList);

                        return;
                    }
                    catch (Exception ex)
                    {
                        lastException = ex;
                        if (conn != null)
                        {
                            conn.Dispose();
                        }
                    }
                }

                throw new InvalidOperationException("无法通过 OLEDB 写入 Excel。最后错误: " + lastException, lastException);
            }
            catch
            {
                // 发生错误时恢复备份
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

        private static string GetExtendedProperties(string excelPath)
        {
            var ext = (Path.GetExtension(excelPath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xls") return "Excel 8.0;HDR=YES;IMEX=1";
            if (ext == ".xlsx") return "Excel 12.0 Xml;HDR=YES;IMEX=1";
            if (ext == ".xlsm") return "Excel 12.0 Macro;HDR=YES;IMEX=1";
            throw new NotSupportedException("不支持的 Excel 扩展名: " + ext);
        }

        private static string ResolveWorksheetName(OleDbConnection conn, string worksheet)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
            {
                var w = worksheet.Trim();
                if (!w.EndsWith("$", StringComparison.Ordinal))
                {
                    w += "$";
                }
                return w;
            }

            // 如果未指定工作表，使用第一个工作表
            var candidates = GetWorksheetCandidates(conn);
            if (candidates.Count > 0)
            {
                return candidates[0];
            }

            throw new InvalidOperationException("Excel 文件中没有找到工作表");
        }

        private static List<string> GetWorksheetCandidates(OleDbConnection conn)
        {
            var list = new List<string>();
            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (schema == null)
            {
                return list;
            }

            foreach (DataRow row in schema.Rows)
            {
                var obj = row["TABLE_NAME"];
                if (obj == null || obj == DBNull.Value) continue;
                var tbl = obj.ToString() ?? string.Empty;
                var idx = tbl.IndexOf('$');
                if (idx < 0) continue;
                var sheet = tbl.Substring(0, idx + 1).Trim('\'');
                if (sheet.Length == 0) continue;
                if (!list.Contains(sheet)) list.Add(sheet);
            }

            return list;
        }

        private static void EnsureColumnsExist(OleDbConnection conn, string sheet)
        {
            // 读取现有列
            var columns = GetExistingColumns(conn, sheet);

            // 检查并添加缺失的列（忽略列已存在的错误）
            if (!columns.Contains(Texts.ColMatched))
            {
                try
                {
                    AddColumn(conn, sheet, Texts.ColMatched, "VARCHAR(10)");
                }
                catch (Exception ex)
                {
                    // 如果是因为列已存在导致的错误，忽略它
                    System.Diagnostics.Debug.WriteLine("添加列 '" + Texts.ColMatched + "' 失败（可能已存在）: " + ex.Message);
                }
            }

            if (!columns.Contains(Texts.ColMatchReason))
            {
                try
                {
                    AddColumn(conn, sheet, Texts.ColMatchReason, "VARCHAR(255)");
                }
                catch (Exception ex)
                {
                    // 如果是因为列已存在导致的错误，忽略它
                    System.Diagnostics.Debug.WriteLine("添加列 '" + Texts.ColMatchReason + "' 失败（可能已存在）: " + ex.Message);
                }
            }
        }

        private static HashSet<string> GetExistingColumns(OleDbConnection conn, string sheet)
        {
            var columns = new HashSet<string>(StringComparer.Ordinal);

            try
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT TOP 1 * FROM [" + sheet + "]";
                var da = new OleDbDataAdapter(cmd);
                var dt = new DataTable();
                da.Fill(dt);

                foreach (DataColumn col in dt.Columns)
                {
                    columns.Add(col.ColumnName);
                }
            }
            catch
            {
                // 忽略错误
            }

            return columns;
        }

        private static void AddColumn(OleDbConnection conn, string sheet, string columnName, string dataType)
        {
            try
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = string.Format(
                    CultureInfo.InvariantCulture,
                    "ALTER TABLE [{0}] ADD [{1}] {2}",
                    sheet,
                    columnName,
                    dataType);
                cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("添加列失败: " + columnName, ex);
            }
        }

        private static void UpdateData(OleDbConnection conn, string sheet, List<PersonMatchStatus> matchStatusList)
        {
            foreach (var status in matchStatusList)
            {
                if (string.IsNullOrWhiteSpace(status.Id))
                    continue;

                try
                {
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = string.Format(
                        CultureInfo.InvariantCulture,
                        "UPDATE [{0}] SET [{1}] = ?, [{2}] = ? WHERE [{3}] = ?",
                        sheet,
                        Texts.ColMatched,
                        Texts.ColMatchReason,
                        Texts.ColId);

                    var matchedValue = status.Matched ? Texts.MatchValueYes : Texts.MatchValueNo;
                    var reasonValue = status.MatchReason ?? string.Empty;

                    cmd.Parameters.Add(new OleDbParameter("?", matchedValue));
                    cmd.Parameters.Add(new OleDbParameter("?", reasonValue));
                    cmd.Parameters.Add(new OleDbParameter("?", status.Id));

                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    // 记录错误但继续处理其他行
                    System.Diagnostics.Debug.WriteLine("更新身份证号 " + status.Id + " 失败: " + ex.Message);
                }
            }
        }
    }
}
