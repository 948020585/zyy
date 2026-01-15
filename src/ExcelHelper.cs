using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace CertPhotoSorter
{
    /// <summary>
    /// Excel OLEDB 公共工具类
    /// </summary>
    internal static class ExcelHelper
    {
        public static readonly string[] Providers = new[]
        {
            "Microsoft.ACE.OLEDB.16.0",
            "Microsoft.ACE.OLEDB.12.0",
            "Microsoft.Jet.OLEDB.4.0"
        };

        /// <summary>
        /// 获取读取模式的扩展属性（IMEX=1 用于混合类型数据）
        /// </summary>
        public static string GetReadExtendedProperties(string excelPath)
        {
            var ext = (Path.GetExtension(excelPath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xls") return "Excel 8.0;HDR=YES;IMEX=1";
            if (ext == ".xlsx") return "Excel 12.0 Xml;HDR=YES;IMEX=1";
            if (ext == ".xlsm") return "Excel 12.0 Macro;HDR=YES;IMEX=1";
            throw new NotSupportedException("Unsupported Excel extension: " + ext);
        }

        /// <summary>
        /// 获取写入模式的扩展属性（不使用 IMEX=1）
        /// </summary>
        public static string GetWriteExtendedProperties(string excelPath)
        {
            var ext = (Path.GetExtension(excelPath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xls") return "Excel 8.0;HDR=YES";
            if (ext == ".xlsx") return "Excel 12.0 Xml;HDR=YES";
            if (ext == ".xlsm") return "Excel 12.0 Macro;HDR=YES";
            throw new NotSupportedException("Unsupported Excel extension: " + ext);
        }

        /// <summary>
        /// 构建连接字符串
        /// </summary>
        public static string BuildConnectionString(string provider, string excelPath, string extProps)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "Provider={0};Data Source={1};Extended Properties='{2}'",
                provider,
                excelPath,
                extProps);
        }

        /// <summary>
        /// 获取工作表候选列表
        /// </summary>
        public static List<string> GetWorksheetCandidates(OleDbConnection conn)
        {
            var list = new List<string>();
            var schema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (schema == null) return list;

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

        /// <summary>
        /// 解析工作表名称（智能选择包含必需列且名称包含"最终"/"报名"的工作表）
        /// </summary>
        public static string ResolveWorksheetName(OleDbConnection conn, string worksheet, bool requireColumns)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
            {
                var w = worksheet.Trim();
                if (!w.EndsWith("$", StringComparison.Ordinal)) w += "$";
                return w;
            }

            var candidates = GetWorksheetCandidates(conn);
            if (candidates.Count == 0)
                throw new InvalidOperationException("Excel file contains no worksheets.");

            if (!requireColumns) return candidates[0];

            string bestSheet = null;
            long bestScore = long.MinValue;

            foreach (var sheet in candidates)
            {
                if (!HasRequiredColumns(conn, sheet)) continue;

                var rowCount = TryGetRowCount(conn, sheet);
                var rank = GetSheetNameRank(sheet);
                var score = (rank * 1000000000L) + rowCount;
                if (score > bestScore)
                {
                    bestScore = score;
                    bestSheet = sheet;
                }
            }

            if (bestSheet != null) return bestSheet;
            throw new InvalidOperationException("No worksheet contains required columns.");
        }

        public static bool HasRequiredColumns(OleDbConnection conn, string sheet)
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT TOP 1 [" + Texts.ColId + "],[" + Texts.ColCert + "] FROM [" + sheet + "]";
                    using (var da = new OleDbDataAdapter(cmd))
                    {
                        var dt = new DataTable();
                        da.Fill(dt);
                        return dt.Columns.Contains(Texts.ColId) && dt.Columns.Contains(Texts.ColCert);
                    }
                }
            }
            catch { return false; }
        }

        public static long TryGetRowCount(OleDbConnection conn, string sheet)
        {
            try
            {
                using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "SELECT COUNT(*) FROM [" + sheet + "]";
                    var obj = cmd.ExecuteScalar();
                    if (obj == null || obj == DBNull.Value) return 0;
                    return Convert.ToInt64(obj, CultureInfo.InvariantCulture);
                }
            }
            catch { return 0; }
        }

        public static int GetSheetNameRank(string sheet)
        {
            var name = sheet ?? string.Empty;
            if (name.EndsWith("$", StringComparison.Ordinal))
                name = name.Substring(0, name.Length - 1);

            var hasFinal = name.IndexOf("\u6700\u7EC8", StringComparison.Ordinal) >= 0;  // 最终
            var hasSignup = name.IndexOf("\u62A5\u540D", StringComparison.Ordinal) >= 0; // 报名

            if (hasFinal && hasSignup) return 3;
            if (hasFinal) return 2;
            if (hasSignup) return 1;
            return 0;
        }
    }
}
