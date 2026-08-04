using System.Linq;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// WordReader 自定义 &lt;table&gt;/&lt;row&gt;/&lt;cell&gt; 标签判定（与 Python table_xml_utils 对齐）。
    /// </summary>
    public static class TableXmlHelper
    {
        /// <summary>
        /// 最外层 table：1 个直接 row，且该 row 下 1 个直接 cell。
        /// </summary>
        public static bool IsSingleCellTableXml(XElement tableElem)
        {
            if (tableElem == null)
            {
                return false;
            }

            var rows = tableElem.Elements("row").ToList();
            if (rows.Count != 1)
            {
                return false;
            }

            var cells = rows[0].Elements("cell").ToList();
            return cells.Count == 1;
        }
    }
}
