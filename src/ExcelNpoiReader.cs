using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using NPOI.SS.UserModel;

namespace CertPhotoSorter
{
    internal static class ExcelNpoiReader
    {
        private const int MaxHeaderScanRows = 30;

        public static List<string> ListWorksheetNames(string excelPath)
        {
            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var workbook = WorkbookFactory.Create(fs);
                try
                {
                    var list = new List<string>();
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        var sheet = workbook.GetSheetAt(i);
                        if (sheet == null) continue;
                        var name = sheet.SheetName ?? string.Empty;
                        if (name.Length == 0) continue;
                        list.Add(name + "$");
                    }
                    return list;
                }
                finally
                {
                    workbook.Close();
                }
            }
        }

        public static DataTable ReadRows(string excelPath, string worksheet, out string worksheetUsed)
        {
            using (var fs = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var workbook = WorkbookFactory.Create(fs);
                try
                {
                    var sheetInfo = ResolveSheetInfo(workbook, worksheet);
                    worksheetUsed = sheetInfo.Sheet.SheetName + "$";
                    return ReadRowsFromSheet(workbook, sheetInfo);
                }
                finally
                {
                    workbook.Close();
                }
            }
        }

        private static SheetInfo ResolveSheetInfo(IWorkbook workbook, string worksheet)
        {
            if (!string.IsNullOrWhiteSpace(worksheet))
            {
                var sheetName = worksheet.Trim();
                if (sheetName.EndsWith("$", StringComparison.Ordinal))
                    sheetName = sheetName.Substring(0, sheetName.Length - 1);

                var sheet = workbook.GetSheet(sheetName);
                if (sheet == null)
                    throw new InvalidOperationException("Worksheet not found: " + worksheet);

                var info = TryBuildSheetInfo(workbook, sheet);
                if (info == null)
                    throw new InvalidOperationException("Worksheet contains no required columns: " + sheetName);

                return info;
            }

            SheetInfo best = null;
            long bestScore = long.MinValue;

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                if (sheet == null) continue;

                var info = TryBuildSheetInfo(workbook, sheet);
                if (info == null) continue;

                var rank = ExcelHelper.GetSheetNameRank(sheet.SheetName + "$");
                var score = (rank * 1000000000L) + info.RowCount;
                if (score > bestScore)
                {
                    bestScore = score;
                    best = info;
                }
            }

            if (best != null) return best;
            throw new InvalidOperationException("No worksheet contains required columns.");
        }

        private static SheetInfo TryBuildSheetInfo(IWorkbook workbook, ISheet sheet)
        {
            int headerRowIndex;
            int idCol;
            int certCol;
            int nameCol;

            if (!TryFindHeader(sheet, workbook, out headerRowIndex, out idCol, out certCol, out nameCol))
                return null;

            var rowCount = CountDataRows(workbook, sheet, headerRowIndex, idCol, certCol, nameCol);
            return new SheetInfo
            {
                Sheet = sheet,
                HeaderRowIndex = headerRowIndex,
                IdCol = idCol,
                CertCol = certCol,
                NameCol = nameCol,
                RowCount = rowCount
            };
        }

        private static bool TryFindHeader(
            ISheet sheet,
            IWorkbook workbook,
            out int headerRowIndex,
            out int idCol,
            out int certCol,
            out int nameCol)
        {
            headerRowIndex = -1;
            idCol = -1;
            certCol = -1;
            nameCol = -1;

            var formatter = new DataFormatter(CultureInfo.InvariantCulture);
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            var start = Math.Max(sheet.FirstRowNum, 0);
            var end = Math.Min(sheet.LastRowNum, start + MaxHeaderScanRows);

            for (int r = start; r <= end; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                int localId = -1;
                int localCert = -1;
                int localName = -1;

                for (int c = row.FirstCellNum; c < row.LastCellNum; c++)
                {
                    if (c < 0) continue;
                    var cell = row.GetCell(c);
                    var text = GetCellText(cell, formatter, evaluator);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    if (localId < 0 && string.Equals(text, Texts.ColId, StringComparison.Ordinal)) localId = c;
                    if (localCert < 0 && string.Equals(text, Texts.ColCert, StringComparison.Ordinal)) localCert = c;
                    if (localName < 0 && string.Equals(text, Texts.ColName, StringComparison.Ordinal)) localName = c;
                }

                if (localId >= 0 && localCert >= 0)
                {
                    headerRowIndex = r;
                    idCol = localId;
                    certCol = localCert;
                    nameCol = localName; // optional
                    return true;
                }
            }

            return false;
        }

        private static long CountDataRows(IWorkbook workbook, ISheet sheet, int headerRowIndex, int idCol, int certCol, int nameCol)
        {
            var formatter = new DataFormatter(CultureInfo.InvariantCulture);
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            long count = 0;
            for (int r = headerRowIndex + 1; r <= sheet.LastRowNum; r++)
            {
                var row = sheet.GetRow(r);
                if (row == null) continue;

                var id = GetCellText(row.GetCell(idCol), formatter, evaluator);
                var cert = GetCellText(row.GetCell(certCol), formatter, evaluator);
                var name = nameCol >= 0 ? GetCellText(row.GetCell(nameCol), formatter, evaluator) : null;

                if (string.IsNullOrWhiteSpace(id) && string.IsNullOrWhiteSpace(cert) && string.IsNullOrWhiteSpace(name))
                    continue;

                count++;
            }
            return count;
        }

        private static DataTable ReadRowsFromSheet(IWorkbook workbook, SheetInfo info)
        {
            var formatter = new DataFormatter(CultureInfo.InvariantCulture);
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();

            var dt = new DataTable();
            dt.Columns.Add(Texts.ColId, typeof(string));
            dt.Columns.Add(Texts.ColCert, typeof(string));
            dt.Columns.Add(Texts.ColName, typeof(string));

            for (int r = info.HeaderRowIndex + 1; r <= info.Sheet.LastRowNum; r++)
            {
                var row = info.Sheet.GetRow(r);
                if (row == null) continue;

                var id = GetCellText(row.GetCell(info.IdCol), formatter, evaluator);
                var cert = GetCellText(row.GetCell(info.CertCol), formatter, evaluator);
                var name = info.NameCol >= 0 ? GetCellText(row.GetCell(info.NameCol), formatter, evaluator) : null;

                if (string.IsNullOrWhiteSpace(id) && string.IsNullOrWhiteSpace(cert) && string.IsNullOrWhiteSpace(name))
                    continue;

                var dr = dt.NewRow();
                dr[Texts.ColId] = id ?? string.Empty;
                dr[Texts.ColCert] = cert ?? string.Empty;
                dr[Texts.ColName] = name ?? string.Empty;
                dt.Rows.Add(dr);
            }

            return dt;
        }

        private static string GetCellText(ICell cell, DataFormatter formatter, IFormulaEvaluator evaluator)
        {
            if (cell == null) return null;
            var text = formatter.FormatCellValue(cell, evaluator);
            if (text == null) return null;
            return text.Trim();
        }

        private sealed class SheetInfo
        {
            public ISheet Sheet;
            public int HeaderRowIndex;
            public int IdCol;
            public int CertCol;
            public int NameCol;
            public long RowCount;
        }
    }
}

