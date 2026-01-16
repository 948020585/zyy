using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;

namespace CertPhotoSorter
{
    internal static class ExcelReader
    {
        public static List<string> ListWorksheetNames(string excelPath, out string providerUsed)
        {
            var extProps = ExcelHelper.GetReadExtendedProperties(excelPath);

            Exception last = null;
            foreach (var provider in ExcelHelper.Providers)
            {
                try
                {
                    var connStr = ExcelHelper.BuildConnectionString(provider, excelPath, extProps);
                    using (var conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        providerUsed = provider;
                        return ExcelHelper.GetWorksheetCandidates(conn);
                    }
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            try
            {
                providerUsed = "NPOI";
                return ExcelNpoiReader.ListWorksheetNames(excelPath);
            }
            catch (Exception npoiEx)
            {
                throw new InvalidOperationException(
                    "Unable to list worksheets via OLEDB or NPOI. OLEDB last error: " + last + "; NPOI error: " + npoiEx,
                    npoiEx);
            }
        }

        public static DataTable ReadRows(string excelPath, string worksheet, out string providerUsed, out string worksheetUsed)
        {
            var extProps = ExcelHelper.GetReadExtendedProperties(excelPath);

            Exception last = null;
            foreach (var provider in ExcelHelper.Providers)
            {
                try
                {
                    var connStr = ExcelHelper.BuildConnectionString(provider, excelPath, extProps);
                    using (var conn = new OleDbConnection(connStr))
                    {
                        conn.Open();
                        var sheet = ExcelHelper.ResolveWorksheetName(conn, worksheet, true);

                        using (var cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "SELECT [" + Texts.ColId + "],[" + Texts.ColCert + "],[" + Texts.ColName + "] FROM [" + sheet + "]";
                            using (var da = new OleDbDataAdapter(cmd))
                            {
                                var dt = new DataTable();
                                da.Fill(dt);
                                providerUsed = provider;
                                worksheetUsed = sheet;
                                return dt;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            try
            {
                providerUsed = "NPOI";
                return ExcelNpoiReader.ReadRows(excelPath, worksheet, out worksheetUsed);
            }
            catch (Exception npoiEx)
            {
                throw new InvalidOperationException(
                    "Unable to read Excel via OLEDB or NPOI. OLEDB last error: " + last + "; NPOI error: " + npoiEx,
                    npoiEx);
            }
        }
    }
}
