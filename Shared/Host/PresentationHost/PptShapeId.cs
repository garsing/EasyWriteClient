using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace WordAddIn1.PresentationHost
{
    /// <summary>约定 ShapeId / section ShapeId 解析（apply / animation 共用）。</summary>
    internal static class PptShapeId
    {
        public static bool TryParseSection(string shapeId, out string slideId)
        {
            slideId = null;
            Match m = Regex.Match(shapeId ?? "", @"^sid(\d+)$", RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                return false;
            }

            slideId = m.Groups[1].Value;
            return true;
        }

        public static bool TryParseShape(string shapeId, out string slideId, out int comId)
        {
            slideId = null;
            comId = 0;
            Match m = Regex.Match(shapeId ?? "", @"^sid(\d+)-s(\d+)$", RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                return false;
            }

            slideId = m.Groups[1].Value;
            return int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out comId);
        }

        public static string FormatShape(string slideId, int comId)
        {
            return "sid" + (slideId ?? "") + "-s" + comId.ToString(CultureInfo.InvariantCulture);
        }
    }
}
