using System.IO;
using System.Text.RegularExpressions;

namespace CertPhotoSorter
{
    /// <summary>
    /// 从文件名提取的姓名和身份证号
    /// </summary>
    internal sealed class NameIdPair
    {
        public string Name;
        public string Id;
    }

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

        /// <summary>
        /// 从文件名中提取姓名和身份证号
        /// 支持格式：张三370102199001011234.jpg 或 370102199001011234张三.jpg
        /// </summary>
        public static NameIdPair ExtractNameAndIdFromFileName(string fileName)
        {
            var baseName = Path.GetFileNameWithoutExtension(fileName) ?? string.Empty;
            var result = new NameIdPair();

            // 优先匹配18位身份证
            var m18 = Id18Regex.Match(baseName);
            if (m18.Success)
            {
                result.Id = m18.Groups[1].Value.ToUpperInvariant();
                result.Name = ExtractNamePart(baseName, m18.Index, m18.Length);
                return result;
            }

            // 尝试17位（自动补校验位）
            var m17 = Id17Regex.Match(baseName);
            if (m17.Success)
            {
                var base17 = m17.Groups[1].Value;
                var chk = ComputeIdCheckDigit(base17);
                if (chk.HasValue)
                {
                    result.Id = base17 + chk.Value;
                    result.Name = ExtractNamePart(baseName, m17.Index, m17.Length);
                    return result;
                }
            }

            // 尝试15位
            var m15 = Id15Regex.Match(baseName);
            if (m15.Success)
            {
                result.Id = m15.Groups[1].Value;
                result.Name = ExtractNamePart(baseName, m15.Index, m15.Length);
                return result;
            }

            return result;
        }

        /// <summary>
        /// 从文件名中提取姓名部分（身份证号前后的非数字部分）
        /// </summary>
        private static string ExtractNamePart(string baseName, int idIndex, int idLength)
        {
            // 身份证号前面的部分
            var before = idIndex > 0 ? baseName.Substring(0, idIndex).Trim() : string.Empty;
            // 身份证号后面的部分
            var afterStart = idIndex + idLength;
            var after = afterStart < baseName.Length ? baseName.Substring(afterStart).Trim() : string.Empty;

            // 优先取非空的部分，如果都有内容则取较短的（通常姓名比其他信息短）
            if (!string.IsNullOrEmpty(before) && !string.IsNullOrEmpty(after))
            {
                return before.Length <= after.Length ? before : after;
            }
            if (!string.IsNullOrEmpty(before)) return before;
            if (!string.IsNullOrEmpty(after)) return after;
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

