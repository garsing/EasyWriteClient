using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    /// <summary>
    /// SVG 体积与安全门禁（硬拒绝危险内容）。
    /// </summary>
    public static class SvgSecurityValidator
    {
        public const int DefaultMaxBytes = 512 * 1024;

        private static readonly Regex EventAttrRegex = new Regex(
            @"\son[a-zA-Z]+\s*=",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex ExternalUrlRegex = new Regex(
            @"(?:href|xlink:href)\s*=\s*[""']\s*(?:https?:|//|javascript:)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public sealed class ValidationResult
        {
            public bool Success { get; set; }
            public string ErrorCode { get; set; }
            public string Error { get; set; }
            public string NormalizedSvg { get; set; }
        }

        public static ValidationResult Validate(string svg, int maxBytes = DefaultMaxBytes)
        {
            if (string.IsNullOrWhiteSpace(svg))
            {
                return Fail("invalid_svg", "svg 为空");
            }

            string trimmed = svg.Trim();
            int byteCount = Encoding.UTF8.GetByteCount(trimmed);
            if (byteCount > maxBytes)
            {
                return Fail("svg_too_large", $"svg 过大（{byteCount} 字节，上限 {maxBytes}）");
            }

            string lower = trimmed.ToLowerInvariant();
            if (lower.Contains("<script") || lower.Contains("</script"))
            {
                return Fail("unsafe_svg", "svg 含 script，已拒绝");
            }

            if (lower.Contains("<foreignobject") || lower.Contains("<iframe"))
            {
                return Fail("unsafe_svg", "svg 含 foreignObject/iframe，已拒绝");
            }

            if (EventAttrRegex.IsMatch(trimmed))
            {
                return Fail("unsafe_svg", "svg 含事件属性，已拒绝");
            }

            if (ExternalUrlRegex.IsMatch(trimmed) || lower.Contains("javascript:"))
            {
                return Fail("unsafe_svg", "svg 含外部或 javascript 链接，已拒绝");
            }

            try
            {
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    IgnoreComments = true,
                    IgnoreWhitespace = true
                };

                using (var reader = XmlReader.Create(new System.IO.StringReader(trimmed), settings))
                {
                    bool foundRoot = false;
                    while (reader.Read())
                    {
                        if (reader.NodeType != XmlNodeType.Element)
                        {
                            continue;
                        }

                        string local = reader.LocalName;
                        if (!foundRoot)
                        {
                            if (!string.Equals(local, "svg", StringComparison.OrdinalIgnoreCase))
                            {
                                return Fail("invalid_svg", "根节点必须是 svg");
                            }

                            foundRoot = true;
                            break;
                        }
                    }

                    if (!foundRoot)
                    {
                        return Fail("invalid_svg", "未找到 svg 根节点");
                    }
                }
            }
            catch (Exception ex)
            {
                return Fail("invalid_svg", $"svg 解析失败: {ex.Message}");
            }

            return new ValidationResult
            {
                Success = true,
                NormalizedSvg = trimmed
            };
        }

        private static ValidationResult Fail(string code, string message)
        {
            return new ValidationResult
            {
                Success = false,
                ErrorCode = code,
                Error = $"{code}: {message}"
            };
        }
    }
}
