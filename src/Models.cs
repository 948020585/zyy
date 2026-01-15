namespace CertPhotoSorter
{
    internal sealed class RunSettings
    {
        public string ExcelPath;
        public string PhotoRoot;
        public string OutputRoot;
        public string Worksheet;
        public bool DryRun;
    }

    internal sealed class RunResult
    {
        public string ProviderUsed;
        public string WorksheetUsed;

        public int ExcelRows;
        public int ExcelPeopleUnique;

        public int PhotoFiles;
        public int MatchedSourceFiles;
        public int MatchedPeople;
        public int MatchedCopies;
        public int UnmatchedNoId;
        public int UnmatchedNotInExcel;

        public string ReportPath;
        public string DetailsCsvPath;
        public string CertSummaryCsvPath;
    }

    internal sealed class OpRow
    {
        public string MatchType;
        public string Cert;
        public string Name;
        public string Id;
        public string Source;
        public string Destination;
    }

    /// <summary>
    /// Excel 中每一行人员的匹配状态
    /// </summary>
    internal sealed class PersonMatchStatus
    {
        public string Id;
        public string Name;
        public string Cert;
        public bool Matched;
        public string MatchReason;
        public int PhotoCount;
    }
}
