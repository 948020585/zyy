namespace CertPhotoSorter
{
    internal static class Texts
    {
        public const string AppTitle = "\u6309\u8BC1\u4E66\u5206\u7167\u7247\u5DE5\u5177";

        public const string DefaultOutputFolder = "\u6309\u8BC1\u4E66\u5206\u7C7B";
        public const string UnmatchedFolder = "\u672A\u5339\u914D";
        public const string EmptyCert = "\u672A\u586B\u5199\u8BC1\u4E66";

        public const string ColId = "\u8EAB\u4EFD\u8BC1\u53F7\u7801";
        public const string ColCert = "\u8BC1\u4E66";
        public const string ColName = "\u59D3\u540D";

        public const string MatchTypeMatched = "\u5339\u914D";
        public const string MatchTypeSkipped = "\u5DF2\u8DF3\u8FC7-\u6587\u4EF6\u5DF2\u5B58\u5728";
        public const string MatchTypeNoId = "\u672A\u5339\u914D-\u65E0\u8EAB\u4EFD\u8BC1\u53F7";
        public const string MatchTypeNotInExcel = "\u672A\u5339\u914D-\u8EAB\u4EFD\u8BC1\u53F7\u4E0D\u5728Excel";

        // Excel 匹配状态列名
        public const string ColMatched = "\u662F\u5426\u5339\u914D";
        public const string ColMatchReason = "\u5339\u914D\u539F\u56E0";

        // Excel 匹配状态值
        public const string MatchValueYes = "\u662F";
        public const string MatchValueNo = "\u5426";
        public const string MatchReasonSuccess = "\u5DF2\u5339\u914D\u5230\u7167\u7247";
        public const string MatchReasonNoPhoto = "\u672A\u5339\u914D\u5230\u7167\u7247";

        public const string CsvHeaderDetails = "\u5339\u914D\u7C7B\u578B,\u8BC1\u4E66,\u59D3\u540D,\u8EAB\u4EFD\u8BC1\u53F7,\u6E90\u6587\u4EF6,\u76EE\u6807\u6587\u4EF6";
        public const string CsvHeaderCertSummary = "\u8BC1\u4E66,Excel\u4EBA\u6570,\u5339\u914D\u5230\u7167\u7247_\u4EBA\u6570,\u5339\u914D\u5230\u7167\u7247_\u6587\u4EF6\u6570,\u672A\u5339\u914D\u7167\u7247\u4EBA\u6570";

        public const string ReportPrefix = "\u8FD0\u884C\u62A5\u544A_";
        public const string DetailsPrefix = "\u660E\u7EC6_";
        public const string CertSummaryPrefix = "\u8BC1\u4E66\u6C47\u603B_";

        public const string UiLabelExcel = "Excel\u6587\u4EF6\uFF1A";
        public const string UiLabelPhotos = "\u7167\u7247\u76EE\u5F55\uFF1A";
        public const string UiLabelOutput = "\u8F93\u51FA\u76EE\u5F55\uFF1A";
        public const string UiLabelWorksheet = "\u5DE5\u4F5C\u8868\uFF08\u53EF\u9009\uFF09\uFF1A";
        public const string UiAutoDetect = "\u81EA\u52A8\u8BC6\u522B";
        public const string UiDryRun = "DryRun\uFF08\u53EA\u751F\u6210\u62A5\u544A\uFF0C\u4E0D\u590D\u5236\uFF09";
        public const string UiChooseFile = "\u9009\u62E9\u6587\u4EF6";
        public const string UiChooseFolder = "\u9009\u62E9\u6587\u4EF6\u5939";
        public const string UiLoadSheets = "\u8BFB\u53D6\u5DE5\u4F5C\u8868";
        public const string UiRun = "\u5F00\u59CB\u6267\u884C";

        public const string MsgNeedExcel = "\u8BF7\u5148\u9009\u62E9\u6709\u6548\u7684Excel\u6587\u4EF6\u3002";
        public const string MsgNeedPhotoRoot = "\u8BF7\u5148\u9009\u62E9\u6709\u6548\u7684\u7167\u7247\u76EE\u5F55\u3002";
        public const string MsgDone = "\u5B8C\u6210\u3002";
        public const string MsgFailed = "\u5931\u8D25\u3002";

        public const string LogReadingExcel = "\u8BFB\u53D6Excel\u2026";
        public const string LogScanningPhotos = "\u626B\u63CF\u7167\u7247\u2026";
        public const string LogWritingReport = "\u5199\u5165\u62A5\u544A\u2026";

        public const string ReportLabelExcel = "Excel\uFF1A";
        public const string ReportLabelWorksheet = "\u5DE5\u4F5C\u8868\uFF1A";
        public const string ReportLabelProvider = "\u63D0\u4F9B\u7A0B\u5E8F\uFF1A";
        public const string ReportLabelPhotoRoot = "\u7167\u7247\u76EE\u5F55\uFF1A";
        public const string ReportLabelOutputRoot = "\u8F93\u51FA\u76EE\u5F55\uFF1A";
        public const string ReportLabelDryRun = "DryRun\uFF1A";
        public const string ReportLabelExcelCounts = "Excel\u884C\u6570\uFF1A";
        public const string ReportLabelPhotoCounts = "\u7167\u7247\u6587\u4EF6\u6570\uFF1A";
        public const string ReportLabelMatched = "\u5339\u914D\uFF1A";
        public const string ReportLabelUnmatchedNoId = "\u672A\u5339\u914D(\u65E0\u8EAB\u4EFD\u8BC1\u53F7)\uFF1A";
        public const string ReportLabelUnmatchedNotInExcel = "\u672A\u5339\u914D(\u8EAB\u4EFD\u8BC1\u53F7\u4E0D\u5728Excel)\uFF1A";
        public const string ReportLabelDetails = "\u660E\u7EC6\uFF1A";
        public const string ReportLabelCertSummary = "\u8BC1\u4E66\u6C47\u603B\uFF1A";

        public const string ReportExcelUniquePeoplePrefix = "\uFF1BExcel\u4EBA\u6570(\u8EAB\u4EFD\u8BC1\u53F7\u53BB\u91CD)\uFF1A";
        public const string ReportUnitZhang = "\u0020\u5F20";
        public const string ReportMatchedTailFormat = "\uFF0C\u8986\u76D6 {0} \u4EBA\uFF08\u590D\u5236\u5230\u8BC1\u4E66\u6587\u4EF6\u5939\u5171 {1} \u4EFD\uFF09";
    }
}
