using System.Linq;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// WordReader 自定义 &lt;table&gt;/&lt;tr&gt;/&lt;td&gt; 标签判定（与 Python table_xml_utils 对齐）。
    /// </summary>
    public static class TableXmlHelper
    {
        /// <summary>
        /// 最外层 table：1 个直接 tr，且该 tr 下 1 个直接 td/th。
        /// </summary>
        public static bool IsSingleCellTableXml(XElement tableElem)
        {
            if (tableElem == null)
            {
                return false;
            }

            var rows = TableMarkupTags.GridRows(tableElem).ToList();
            if (rows.Count != 1)
            {
                return false;
            }

            var cells = TableMarkupTags.GridCells(rows[0]).ToList();
            return cells.Count == 1;
        }
    }
}
