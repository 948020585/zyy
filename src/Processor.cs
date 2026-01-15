using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Text;

namespace CertPhotoSorter
{
    internal static class Processor
    {
        private static readonly HashSet<string> ImageExts = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".jpg",
            ".jpeg",
            ".png",
            ".bmp"
        };

        private static readonly char[] InvalidFileNameChars = Path.GetInvalidFileNameChars();

        public static RunResult Execute(RunSettings settings, Action<string> log, Action<int> progress)
        {
            if (settings == null) throw new ArgumentNullException("settings");
            if (string.IsNullOrWhiteSpace(settings.ExcelPath)) throw new ArgumentException("ExcelPath is required.");
            if (string.IsNullOrWhiteSpace(settings.PhotoRoot)) throw new ArgumentException("PhotoRoot is required.");

            settings.ExcelPath = Path.GetFullPath(settings.ExcelPath);
            settings.PhotoRoot = Path.GetFullPath(settings.PhotoRoot);
            settings.OutputRoot = string.IsNullOrWhiteSpace(settings.OutputRoot)
                ? Path.Combine(Path.GetDirectoryName(settings.ExcelPath) ?? Environment.CurrentDirectory, Texts.DefaultOutputFolder)
                : Path.GetFullPath(settings.OutputRoot);

            if (!File.Exists(settings.ExcelPath)) throw new FileNotFoundException("Excel file not found.", settings.ExcelPath);
            if (!Directory.Exists(settings.PhotoRoot)) throw new DirectoryNotFoundException("Photo root not found: " + settings.PhotoRoot);

            Directory.CreateDirectory(settings.OutputRoot);
            var unmatchedDir = Path.Combine(settings.OutputRoot, Texts.UnmatchedFolder);
            Directory.CreateDirectory(unmatchedDir);

            if (log != null) log(Texts.LogReadingExcel);

            string providerUsed;
            string worksheetUsed;
            var excelRows = ExcelReader.ReadRows(settings.ExcelPath, settings.Worksheet, out providerUsed, out worksheetUsed);

            var idToCerts = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var idToName = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var idToCertAndName = new Dictionary<string, List<PersonMatchStatus>>(StringComparer.OrdinalIgnoreCase);
            var excelTotalsByCert = new Dictionary<string, int>(StringComparer.Ordinal);

            foreach (DataRow row in excelRows.Rows)
            {
                var id = Normalize(row[Texts.ColId]);
                if (string.IsNullOrWhiteSpace(id)) continue;
                id = id.ToUpperInvariant();

                var cert = Normalize(row[Texts.ColCert]);
                if (string.IsNullOrWhiteSpace(cert)) cert = Texts.EmptyCert;

                var name = Normalize(row[Texts.ColName]);

                HashSet<string> certs;
                if (!idToCerts.TryGetValue(id, out certs))
                {
                    certs = new HashSet<string>(StringComparer.Ordinal);
                    idToCerts[id] = certs;
                }
                certs.Add(cert);

                if (!string.IsNullOrWhiteSpace(name) && !idToName.ContainsKey(id))
                {
                    idToName[id] = name;
                }

                // 初始化匹配状态（默认未匹配）
                List<PersonMatchStatus> statusList;
                if (!idToCertAndName.TryGetValue(id, out statusList))
                {
                    statusList = new List<PersonMatchStatus>();
                    idToCertAndName[id] = statusList;
                }
                statusList.Add(new PersonMatchStatus
                {
                    Id = id,
                    Name = name,
                    Cert = cert,
                    Matched = false,
                    MatchReason = Texts.MatchReasonNoPhoto,
                    PhotoCount = 0
                });

                int total;
                excelTotalsByCert.TryGetValue(cert, out total);
                excelTotalsByCert[cert] = total + 1;
            }

            if (log != null)
            {
                log("ExcelRows: " + excelRows.Rows.Count + "; UniqueIds: " + idToCerts.Count);
                log(Texts.LogScanningPhotos);
            }

            var photoFiles = GetPhotoFiles(settings.PhotoRoot);
            var ops = new List<OpRow>(capacity: photoFiles.Count);

            var matchedUniqueIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var matchedSourceFiles = 0;
            var matchedCopies = 0;
            var unmatchedNoId = 0;
            var unmatchedNotInExcel = 0;

            for (int i = 0; i < photoFiles.Count; i++)
            {
                if (progress != null)
                {
                    var percent = (int)Math.Round((i + 1) * 100.0 / Math.Max(1, photoFiles.Count));
                    progress(Math.Max(0, Math.Min(100, percent)));
                }

                var src = photoFiles[i];
                var fileName = Path.GetFileName(src) ?? string.Empty;

                var id = IdUtils.ExtractIdFromFileName(fileName);
                if (string.IsNullOrWhiteSpace(id))
                {
                    unmatchedNoId++;
                    var dest = Path.Combine(unmatchedDir, fileName);
                    var nameFromFile = Path.GetFileNameWithoutExtension(fileName);

                    // 检查文件是否已存在，如存在则跳过
                    bool shouldCopy = true;
                    if (File.Exists(dest))
                    {
                        var existingInfo = new FileInfo(dest);
                        var sourceInfo = new FileInfo(src);
                        if (existingInfo.Length == sourceInfo.Length)
                        {
                            shouldCopy = false;
                            ops.Add(new OpRow
                            {
                                MatchType = Texts.MatchTypeSkipped,
                                Cert = null,
                                Name = nameFromFile,
                                Id = null,
                                Source = src,
                                Destination = dest
                            });
                        }
                        else
                        {
                            dest = GetUniqueDestinationPath(dest);
                        }
                    }

                    if (shouldCopy)
                    {
                        ops.Add(new OpRow
                        {
                            MatchType = Texts.MatchTypeNoId,
                            Cert = null,
                            Name = nameFromFile,
                            Id = null,
                            Source = src,
                            Destination = dest
                        });
                        if (!settings.DryRun)
                        {
                            File.Copy(src, dest, true);
                        }
                    }
                    continue;
                }

                HashSet<string> certs2;
                if (!idToCerts.TryGetValue(id, out certs2))
                {
                    unmatchedNotInExcel++;
                    var dest = Path.Combine(unmatchedDir, fileName);
                    var nameFromFile = Path.GetFileNameWithoutExtension(fileName);

                    // 检查文件是否已存在，如存在则跳过
                    bool shouldCopy = true;
                    if (File.Exists(dest))
                    {
                        var existingInfo = new FileInfo(dest);
                        var sourceInfo = new FileInfo(src);
                        if (existingInfo.Length == sourceInfo.Length)
                        {
                            shouldCopy = false;
                            ops.Add(new OpRow
                            {
                                MatchType = Texts.MatchTypeSkipped,
                                Cert = null,
                                Name = nameFromFile,
                                Id = id,
                                Source = src,
                                Destination = dest
                            });
                        }
                        else
                        {
                            dest = GetUniqueDestinationPath(dest);
                        }
                    }

                    if (shouldCopy)
                    {
                        ops.Add(new OpRow
                        {
                            MatchType = Texts.MatchTypeNotInExcel,
                            Cert = null,
                            Name = nameFromFile,
                            Id = id,
                            Source = src,
                            Destination = dest
                        });
                        if (!settings.DryRun)
                        {
                            File.Copy(src, dest, true);
                        }
                    }
                    continue;
                }

                matchedSourceFiles++;
                matchedUniqueIds.Add(id);

                string name2;
                idToName.TryGetValue(id, out name2);

                // 更新匹配状态
                List<PersonMatchStatus> statusList;
                if (idToCertAndName.TryGetValue(id, out statusList))
                {
                    foreach (var status in statusList)
                    {
                        status.Matched = true;
                        status.MatchReason = Texts.MatchReasonSuccess;
                        status.PhotoCount++;
                    }
                }

                foreach (var cert in certs2)
                {
                    var safeCert = ConvertToSafeFolderName(cert);
                    var destDir = Path.Combine(settings.OutputRoot, safeCert);
                    var dest = Path.Combine(destDir, fileName);

                    // 检查目标文件是否已存在（文件名+大小都相同则跳过）
                    bool shouldSkip = false;
                    if (File.Exists(dest))
                    {
                        var existingInfo = new FileInfo(dest);
                        var sourceInfo = new FileInfo(src);
                        if (existingInfo.Length == sourceInfo.Length)
                        {
                            shouldSkip = true;
                            ops.Add(new OpRow
                            {
                                MatchType = Texts.MatchTypeSkipped,
                                Cert = cert,
                                Name = name2,
                                Id = id,
                                Source = src,
                                Destination = dest
                            });
                        }
                    }

                    if (!shouldSkip)
                    {
                        ops.Add(new OpRow
                        {
                            MatchType = Texts.MatchTypeMatched,
                            Cert = cert,
                            Name = name2,
                            Id = id,
                            Source = src,
                            Destination = dest
                        });

                        matchedCopies++;
                        if (!settings.DryRun)
                        {
                            Directory.CreateDirectory(destDir);
                            File.Copy(src, dest, true);
                        }
                    }
                }
            }

            if (progress != null) progress(100);

            if (log != null) log(Texts.LogWritingReport);

            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture);
            var reportPath = Path.Combine(settings.OutputRoot, Texts.ReportPrefix + timestamp + ".txt");
            var detailsCsvPath = Path.Combine(settings.OutputRoot, Texts.DetailsPrefix + timestamp + ".csv");
            var certSummaryCsvPath = Path.Combine(settings.OutputRoot, Texts.CertSummaryPrefix + timestamp + ".csv");

            WriteDetailsCsv(detailsCsvPath, ops);
            WriteCertSummaryCsv(certSummaryCsvPath, excelTotalsByCert, ops);
            WriteReport(
                reportPath,
                settings,
                providerUsed,
                worksheetUsed,
                excelRows.Rows.Count,
                idToCerts.Count,
                photoFiles.Count,
                matchedSourceFiles,
                matchedUniqueIds.Count,
                matchedCopies,
                unmatchedNoId,
                unmatchedNotInExcel,
                detailsCsvPath,
                certSummaryCsvPath);

            // 更新原始Excel文件中的匹配状态
            if (!settings.DryRun)
            {
                if (log != null) log("正在更新Excel文件...");

                var allMatchStatus = new List<PersonMatchStatus>();
                foreach (var kvp in idToCertAndName)
                {
                    allMatchStatus.AddRange(kvp.Value);
                }

                try
                {
                    ExcelWriter.UpdateMatchStatus(settings.ExcelPath, settings.Worksheet, allMatchStatus);
                    if (log != null) log("Excel文件已更新");
                }
                catch (Exception ex)
                {
                    if (log != null) log("Excel文件更新失败: " + ex.Message);
                }
            }

            if (log != null) log(Texts.MsgDone);

            return new RunResult
            {
                ProviderUsed = providerUsed,
                WorksheetUsed = worksheetUsed,
                ExcelRows = excelRows.Rows.Count,
                ExcelPeopleUnique = idToCerts.Count,
                PhotoFiles = photoFiles.Count,
                MatchedSourceFiles = matchedSourceFiles,
                MatchedPeople = matchedUniqueIds.Count,
                MatchedCopies = matchedCopies,
                UnmatchedNoId = unmatchedNoId,
                UnmatchedNotInExcel = unmatchedNotInExcel,
                ReportPath = reportPath,
                DetailsCsvPath = detailsCsvPath,
                CertSummaryCsvPath = certSummaryCsvPath
            };
        }

        private static string Normalize(object value)
        {
            if (value == null || value == DBNull.Value) return null;
            return (value.ToString() ?? string.Empty).Trim();
        }

        private static List<string> GetPhotoFiles(string photoRoot)
        {
            var list = new List<string>();
            foreach (var path in Directory.EnumerateFiles(photoRoot, "*.*", SearchOption.AllDirectories))
            {
                var ext = Path.GetExtension(path);
                if (ImageExts.Contains(ext))
                {
                    list.Add(path);
                }
            }
            return list;
        }

        private static string ConvertToSafeFolderName(string name)
        {
            var result = (name ?? string.Empty).Trim();
            if (result.Length == 0) return Texts.EmptyCert;

            for (int i = 0; i < InvalidFileNameChars.Length; i++)
            {
                result = result.Replace(InvalidFileNameChars[i], '_');
            }

            result = result.Trim().TrimEnd('.', ' ');
            if (result.Length == 0) return Texts.EmptyCert;
            return result;
        }

        private static string GetUniqueDestinationPath(string path)
        {
            if (!File.Exists(path)) return path;

            var dir = Path.GetDirectoryName(path) ?? string.Empty;
            var baseName = Path.GetFileNameWithoutExtension(path) ?? string.Empty;
            var ext = Path.GetExtension(path) ?? string.Empty;

            for (int i = 2; i < 10000; i++)
            {
                var candidate = Path.Combine(dir, string.Format(CultureInfo.InvariantCulture, "{0}_{1}{2}", baseName, i, ext));
                if (!File.Exists(candidate)) return candidate;
            }

            throw new IOException("Unable to generate unique file name: " + path);
        }

        private static void WriteDetailsCsv(string path, List<OpRow> ops)
        {
            using (var sw = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                sw.WriteLine(Texts.CsvHeaderDetails);
                for (int i = 0; i < ops.Count; i++)
                {
                    var op = ops[i];
                    sw.WriteLine(
                        CsvUtils.Escape(op.MatchType) + "," +
                        CsvUtils.Escape(op.Cert) + "," +
                        CsvUtils.Escape(op.Name) + "," +
                        CsvUtils.Escape(op.Id) + "," +
                        CsvUtils.Escape(op.Source) + "," +
                        CsvUtils.Escape(op.Destination));
                }
            }
        }

        private static void WriteCertSummaryCsv(string path, Dictionary<string, int> excelTotalsByCert, List<OpRow> ops)
        {
            var matchedIdsByCert = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
            var matchedFilesByCert = new Dictionary<string, int>(StringComparer.Ordinal);

            for (int i = 0; i < ops.Count; i++)
            {
                var op = ops[i];
                if (!string.Equals(op.MatchType, Texts.MatchTypeMatched, StringComparison.Ordinal)) continue;
                if (string.IsNullOrWhiteSpace(op.Cert) || string.IsNullOrWhiteSpace(op.Id)) continue;

                HashSet<string> ids;
                if (!matchedIdsByCert.TryGetValue(op.Cert, out ids))
                {
                    ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    matchedIdsByCert[op.Cert] = ids;
                }
                ids.Add(op.Id);

                int c;
                matchedFilesByCert.TryGetValue(op.Cert, out c);
                matchedFilesByCert[op.Cert] = c + 1;
            }

            var certKeys = new List<string>(excelTotalsByCert.Keys);
            certKeys.Sort(StringComparer.Ordinal);

            using (var sw = new StreamWriter(path, false, new UTF8Encoding(true)))
            {
                sw.WriteLine(Texts.CsvHeaderCertSummary);

                for (int i = 0; i < certKeys.Count; i++)
                {
                    var cert = certKeys[i];
                    var total = excelTotalsByCert[cert];

                    HashSet<string> ids;
                    matchedIdsByCert.TryGetValue(cert, out ids);
                    var matchedPeople = ids == null ? 0 : ids.Count;

                    int files;
                    matchedFilesByCert.TryGetValue(cert, out files);

                    sw.WriteLine(
                        CsvUtils.Escape(cert) + "," +
                        total.ToString(CultureInfo.InvariantCulture) + "," +
                        matchedPeople.ToString(CultureInfo.InvariantCulture) + "," +
                        files.ToString(CultureInfo.InvariantCulture) + "," +
                        (total - matchedPeople).ToString(CultureInfo.InvariantCulture));
                }
            }
        }

        private static void WriteReport(
            string path,
            RunSettings settings,
            string providerUsed,
            string sheetUsed,
            int excelRows,
            int excelPeopleUnique,
            int photoFiles,
            int matchedSourceFiles,
            int matchedPeople,
            int matchedCopies,
            int unmatchedNoId,
            int unmatchedNotInExcel,
            string detailsCsv,
            string certSummaryCsv)
        {
            var lines = new[]
            {
                Texts.ReportLabelExcel + settings.ExcelPath,
                Texts.ReportLabelWorksheet + sheetUsed,
                Texts.ReportLabelProvider + providerUsed,
                Texts.ReportLabelPhotoRoot + settings.PhotoRoot,
                Texts.ReportLabelOutputRoot + settings.OutputRoot,
                Texts.ReportLabelDryRun + settings.DryRun,
                "",
                Texts.ReportLabelExcelCounts + excelRows + Texts.ReportExcelUniquePeoplePrefix + excelPeopleUnique,
                Texts.ReportLabelPhotoCounts + photoFiles,
                Texts.ReportLabelMatched + matchedSourceFiles + Texts.ReportUnitZhang + string.Format(CultureInfo.InvariantCulture, Texts.ReportMatchedTailFormat, matchedPeople, matchedCopies),
                Texts.ReportLabelUnmatchedNoId + unmatchedNoId + Texts.ReportUnitZhang,
                Texts.ReportLabelUnmatchedNotInExcel + unmatchedNotInExcel + Texts.ReportUnitZhang,
                "",
                Texts.ReportLabelDetails + detailsCsv,
                Texts.ReportLabelCertSummary + certSummaryCsv
            };

            File.WriteAllLines(path, lines, new UTF8Encoding(true));
        }
    }
}
