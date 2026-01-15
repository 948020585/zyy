using System.IO;
using System.Text.RegularExpressions;

namespace CertPhotoSorter
{
    internal static class IdUtils
    {
        private static readonly Regex Id18Regex = new Regex("([0-9]{17}[0-9Xx])", RegexOptions.Compiled);
        private static readonly Regex Id17Regex = new Regex("([0-9]{17})", RegexOptions.Compiled);
        private static readonly Regex Id15Regex = new Regex("([0-9]{15})", RegexOptions.Compiled);

        public static string ExtractIdFromFileName(string fileName)
        {
            var baseName = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;

            var m18 = Id18Regex.Match(baseName);
            if (m18.Success)
            {
                return m18.Groups[1].Value.ToUpperInvariant();
            }

            var m17 = Id17Regex.Match(baseName);
            if (m17.Success)
            {
                var base17 = m17.Groups[1].Value;
                var chk = ComputeIdCheckDigit(base17);
                if (chk.HasValue)
                {
                    return base17 + chk.Value;
                }
            }

            var m15 = Id15Regex.Match(baseName);
            if (m15.Success)
            {
                return m15.Groups[1].Value;
            }

            return null;
        }

        private static char? ComputeIdCheckDigit(string base17)
        {
            if (string.IsNullOrWhiteSpace(base17) || base17.Length != 17)
            {
                return null;
            }

            for (int i = 0; i < base17.Length; i++)
            {
                var ch = base17[i];
                if (ch < '0' || ch > '9')
                {
                    return null;
                }
            }

            var weights = new[] { 7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2 };
            var map = new[] { '1', '0', 'X', '9', '8', '7', '6', '5', '4', '3', '2' };

            var sum = 0;
            for (int i = 0; i < 17; i++)
            {
                sum += (base17[i] - '0') * weights[i];
            }

            return map[sum % 11];
        }
    }
}

