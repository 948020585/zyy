using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace CertPhotoSorter
{
    internal static class CliRunner
    {
        public static int Run(string[] args)
        {
            var settings = new RunSettings();
            string logPath = null;

            for (int i = 0; i < args.Length; i++)
            {
                var a = args[i] ?? string.Empty;
                if (a.Equals("--help", StringComparison.OrdinalIgnoreCase) || a.Equals("-h", StringComparison.OrdinalIgnoreCase))
                {
                    PrintHelp();
                    return 0;
                }

                if (a.Equals("--excel", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    settings.ExcelPath = args[++i];
                    continue;
                }

                if (a.Equals("--photos", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    settings.PhotoRoot = args[++i];
                    continue;
                }

                if (a.Equals("--out", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    settings.OutputRoot = args[++i];
                    continue;
                }

                if (a.Equals("--sheet", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    settings.Worksheet = args[++i];
                    continue;
                }

                if (a.Equals("--dry-run", StringComparison.OrdinalIgnoreCase))
                {
                    settings.DryRun = true;
                    continue;
                }

                if (a.Equals("--match-mode", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    var mode = args[++i];
                    if (mode.Equals("name-id", StringComparison.OrdinalIgnoreCase) || mode.Equals("nameid", StringComparison.OrdinalIgnoreCase))
                    {
                        settings.MatchMode = MatchMode.NameAndId;
                    }
                    continue;
                }

                if (a.Equals("--log", StringComparison.OrdinalIgnoreCase) && i + 1 < args.Length)
                {
                    logPath = args[++i];
                    continue;
                }
            }

            if (string.IsNullOrWhiteSpace(settings.ExcelPath) || string.IsNullOrWhiteSpace(settings.PhotoRoot))
            {
                PrintHelp();
                return 2;
            }

            settings.ExcelPath = Path.GetFullPath(settings.ExcelPath);
            settings.PhotoRoot = Path.GetFullPath(settings.PhotoRoot);
            if (!string.IsNullOrWhiteSpace(settings.OutputRoot))
            {
                settings.OutputRoot = Path.GetFullPath(settings.OutputRoot);
            }
            else
            {
                settings.OutputRoot = Path.Combine(Path.GetDirectoryName(settings.ExcelPath) ?? Environment.CurrentDirectory, Texts.DefaultOutputFolder);
            }

            Directory.CreateDirectory(settings.OutputRoot);

            if (string.IsNullOrWhiteSpace(logPath))
            {
                logPath = Path.Combine(settings.OutputRoot, "EXE_RunLog_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".txt");
            }

            var lines = new List<string>();
            Action<string> log = msg =>
            {
                var line = DateTime.Now.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + " " + msg;
                lines.Add(line);
                Console.WriteLine(line);
            };

            try
            {
                var result = Processor.Execute(settings, log, null);
                lines.Add("Report: " + result.ReportPath);
                File.WriteAllLines(logPath, lines.ToArray(), new UTF8Encoding(true));
                return 0;
            }
            catch (Exception ex)
            {
                lines.Add("ERROR: " + ex);
                try
                {
                    File.WriteAllLines(logPath, lines.ToArray(), new UTF8Encoding(true));
                }
                catch
                {
                    // ignore
                }

                Console.Error.WriteLine(ex);
                return 1;
            }
        }

        private static void PrintHelp()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  CertPhotoSorter.exe --excel <file.xls/xlsx/xlsm> --photos <photoRoot> [--out <outputRoot>] [--sheet <worksheet>] [--match-mode <id|name-id>] [--dry-run] [--log <logPath>]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --match-mode id       Match by ID only (default)");
            Console.WriteLine("  --match-mode name-id  Match by name + ID");
        }
    }
}

