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

        /// <summary>
        /// 仅解析页内 Shape.Id：接受 <c>sid{SlideID}-s{n}</c> 或短写 <c>s{n}</c>。
        /// apply 以工具参数 slide_id 定页，不要求 ShapeId 内 SlideID 与目标页一致。
        /// </summary>
        public static bool TryParseShapeComId(string shapeId, out int comId)
        {
            comId = 0;
            if (TryParseShape(shapeId, out _, out comId))
            {
                return true;
            }

            Match m = Regex.Match(shapeId ?? "", @"^s(\d+)$", RegexOptions.IgnoreCase);
            if (!m.Success)
            {
                return false;
            }

            return int.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out comId);
        }

        public static string FormatShape(string slideId, int comId)
        {
            return "sid" + (slideId ?? "") + "-s" + comId.ToString(CultureInfo.InvariantCulture);
        }
    }
}
