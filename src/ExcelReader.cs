using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace CertPhotoSorter
{
    internal static class ExcelReader
    {
        public static List<string> ListWorksheetNames(string excelPath, out string providerUsed)
        {
            var extProps = GetExtendedProperties(excelPath);
            var providers = new[]
            {
                "Microsoft.ACE.OLEDB.16.0",
                "Microsoft.ACE.OLEDB.12.0",
                "Microsoft.Jet.OLEDB.4.0"
            };

            Exception last = null;
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

                    providerUsed = provider;
                    return GetWorksheetCandidates(conn);
                }
                catch (Exception ex)
                {
                    last = ex;
                    if (conn != null)
                    {
                        conn.Dispose();
                    }
                }
            }

            throw new InvalidOperationException("Unable to list worksheets via OLEDB. Last error: " + last, last);
        }

        public static DataTable ReadRows(string excelPath, string worksheet, out string providerUsed, out string worksheetUsed)
        {
            var extProps = GetExtendedProperties(excelPath);
            var providers = new[]
            {
                "Microsoft.ACE.OLEDB.16.0",
                "Microsoft.ACE.OLEDB.12.0",
                "Microsoft.Jet.OLEDB.4.0"
            };

            Exception last = null;
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

                    var sheet = ResolveWorksheetName(conn, worksheet);
                    var cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT [" + Texts.ColId + "],[" + Texts.ColCert + "],[" + Texts.ColName + "] FROM [" + sheet + "]";

                    var da = new OleDbDataAdapter(cmd);
                    var dt = new DataTable();
                    da.Fill(dt);

                    providerUsed = provider;
                    worksheetUsed = sheet;
                    return dt;
                }
                catch (Exception ex)
                {
                    last = ex;
                    if (conn != null)
                    {
                        conn.Dispose();
                    }
                }
            }

            throw new InvalidOperationException("Unable to read Excel via OLEDB. Last error: " + last, last);
        }

        private static string GetExtendedProperties(string excelPath)
        {
            var ext = (Path.GetExtension(excelPath) ?? string.Empty).ToLowerInvariant();
            if (ext == ".xls") return "Excel 8.0;HDR=YES;IMEX=1";
            if (ext == ".xlsx") return "Excel 12.0 Xml;HDR=YES;IMEX=1";
            if (ext == ".xlsm") return "Excel 12.0 Macro;HDR=YES;IMEX=1";
            throw new NotSupportedException("Unsupported Excel extension: " + ext);
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

            var candidates = GetWorksheetCandidates(conn);
            string bestSheet = null;
            long bestScore = long.MinValue;

            for (int i = 0; i < candidates.Count; i++)
            {
                var sheet = candidates[i];
                if (!HasRequiredColumns(conn, sheet))
                {
                    continue;
                }

                var rowCount = TryGetRowCount(conn, sheet);
                var rank = GetSheetNameRank(sheet);
                var score = (rank * 1000000000L) + rowCount;
                if (score > bestScore)
                {
                    bestScore = score;
                    bestSheet = sheet;
                }
            }

            if (bestSheet != null)
            {
                return bestSheet;
            }

            throw new InvalidOperationException("No worksheet contains required columns.");
        }

        private static bool HasRequiredColumns(OleDbConnection conn, string sheet)
        {
            try
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT TOP 1 [" + Texts.ColId + "],[" + Texts.ColCert + "] FROM [" + sheet + "]";
                var da = new OleDbDataAdapter(cmd);
                var dt = new DataTable();
                da.Fill(dt);
                return dt.Columns.Contains(Texts.ColId) && dt.Columns.Contains(Texts.ColCert);
            }
            catch
            {
                return false;
            }
        }

        private static long TryGetRowCount(OleDbConnection conn, string sheet)
        {
            try
            {
                var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT COUNT(*) FROM [" + sheet + "]";
                var obj = cmd.ExecuteScalar();
                if (obj == null || obj == DBNull.Value)
                {
                    return 0;
                }
                return Convert.ToInt64(obj, CultureInfo.InvariantCulture);
            }
            catch
            {
                return 0;
            }
        }

        private static int GetSheetNameRank(string sheet)
        {
            var name = sheet ?? string.Empty;
            if (name.EndsWith("$", StringComparison.Ordinal))
            {
                name = name.Substring(0, name.Length - 1);
            }

            var hasFinal = name.IndexOf("最终", StringComparison.Ordinal) >= 0;
            var hasSignup = name.IndexOf("报名", StringComparison.Ordinal) >= 0;

            if (hasFinal && hasSignup) return 3;
            if (hasFinal) return 2;
            if (hasSignup) return 1;
            return 0;
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
    }
}
