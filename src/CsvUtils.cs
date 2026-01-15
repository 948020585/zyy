using System;

namespace CertPhotoSorter
{
    internal static class CsvUtils
    {
        public static string Escape(string value)
        {
            if (value == null)
            {
                return "\"\"";
            }

            var v = value;
            var needQuotes = v.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0;
            v = v.Replace("\"", "\"\"");
            return needQuotes ? "\"" + v + "\"" : v;
        }
    }
}

