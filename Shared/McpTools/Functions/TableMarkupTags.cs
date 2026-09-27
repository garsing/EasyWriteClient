using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 表格展示 / TableConfig Data 的行格标签。只换名，不改规则。
    /// </summary>
    public static class TableMarkupTags
    {
        public const string Tr = "tr";
        public const string Td = "td";
        public const string Th = "th";

        public const string LegacyRow = "row";
        public const string LegacyCell = "cell";
        public const string LegacyRowPascal = "Row";
        public const string LegacyCellPascal = "Cell";

        public const string OpenCellPattern = @"<(?:td|th)(?:\s+[^>]*)?>";
        public const string CloseCellPattern = @"</(?:td|th)\s*>";
        public const string OpenRowPattern = @"<tr(?:\s+[^>]*)?>";
        public const string CloseRowPattern = @"</tr\s*>";

        public const string LegacyGridError = "表格 Data 请用 <tr><td>，不要写 <Row><Cell>";

        public static bool IsGridRow(XElement element)
        {
            return element != null && element.Name.LocalName == Tr;
        }

        public static bool IsGridCell(XElement element)
        {
            if (element == null)
            {
                return false;
            }

            string name = element.Name.LocalName;
            return name == Td || name == Th;
        }

        public static bool IsLegacyGridName(string localName)
        {
            if (string.IsNullOrEmpty(localName))
            {
                return false;
            }

            return localName == LegacyRow
                || localName == LegacyCell
                || localName == LegacyRowPascal
                || localName == LegacyCellPascal;
        }

        public static IEnumerable<XElement> GridRows(XElement dataElement)
        {
            if (dataElement == null)
            {
                return Enumerable.Empty<XElement>();
            }

            return dataElement.Elements().Where(IsGridRow);
        }

        public static IEnumerable<XElement> GridCells(XElement rowElement)
        {
            if (rowElement == null)
            {
                return Enumerable.Empty<XElement>();
            }

            return rowElement.Elements().Where(IsGridCell);
        }

        /// <summary>Data 下出现旧行/格标签则失败。Properties/Rows、图表 Rows 不走这里。</summary>
        public static bool TryGetLegacyGridError(XElement dataElement, out string error)
        {
            error = null;
            if (dataElement == null)
            {
                return false;
            }

            foreach (XElement child in dataElement.DescendantsAndSelf())
            {
                if (IsLegacyGridName(child.Name.LocalName))
                {
                    error = LegacyGridError;
                    return true;
                }
            }

            return false;
        }

        public static XElement NewTd(params object[] content)
        {
            return content == null || content.Length == 0
                ? new XElement(Td)
                : new XElement(Td, content);
        }

        public static XElement NewTr(params object[] content)
        {
            return content == null || content.Length == 0
                ? new XElement(Tr)
                : new XElement(Tr, content);
        }

        public static string StripCellTagsRegex()
        {
            return @"<(?:td|th)(?:\s+[^>]*)?>|</(?:td|th)\s*>";
        }

        public static string StructureTagRegex()
        {
            return @"<image(?:\s+[^>]*)?/>|<table(?:\s+[^>]*)?>|<tr(?:\s+[^>]*)?>|<(?:td|th)(?:\s+[^>]*)?>|</table\s*>|</tr\s*>|</(?:td|th)\s*>";
        }

        public static int FindOpenCell(string text, int startPos)
        {
            return FindRegex(text, OpenCellPattern, startPos);
        }

        public static int FindCloseCell(string text, int startPos)
        {
            return FindRegex(text, CloseCellPattern, startPos);
        }

        private static int FindRegex(string text, string pattern, int startPos)
        {
            if (text == null || startPos < 0 || startPos >= text.Length)
            {
                return -1;
            }

            Match match = Regex.Match(text.Substring(startPos), pattern, RegexOptions.IgnoreCase);
            return match.Success ? startPos + match.Index : -1;
        }
    }
}
