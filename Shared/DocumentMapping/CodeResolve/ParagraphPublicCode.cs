using System.Text.RegularExpressions;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// 公开段落码：独一份光杆 <c>P_xxxxx</c>；撞码 <c>P_xxxxx_N</c>（1-based）。
    /// 连字符 <c>P_xxxxx-1</c> 不认。
    /// </summary>
    public static class ParagraphPublicCode
    {
        private static readonly Regex Pattern = new Regex(
            @"^P_([0-9a-zA-Z]{5})(?:_([1-9][0-9]*))?$",
            RegexOptions.Compiled);

        public sealed class Parsed
        {
            public string PublicCode { get; set; }
            public string BaseCode { get; set; }
            public int Occurrence { get; set; }
            public bool RequiresUnique { get; set; }
        }

        public static bool TryParse(string code, out Parsed parsed)
        {
            parsed = null;
            if (string.IsNullOrEmpty(code))
            {
                return false;
            }

            Match m = Pattern.Match(code.Trim());
            if (!m.Success)
            {
                return false;
            }

            string baseCode = "P_" + m.Groups[1].Value;
            bool hasN = m.Groups[2].Success;
            int n = 1;
            if (hasN && !int.TryParse(m.Groups[2].Value, out n))
            {
                return false;
            }

            parsed = new Parsed
            {
                PublicCode = code.Trim(),
                BaseCode = baseCode,
                Occurrence = n,
                RequiresUnique = !hasN
            };
            return true;
        }

        public static bool TryGetBaseCode(string code, out string baseCode)
        {
            if (TryParse(code, out Parsed parsed))
            {
                baseCode = parsed.BaseCode;
                return true;
            }

            baseCode = null;
            return false;
        }
    }
}
