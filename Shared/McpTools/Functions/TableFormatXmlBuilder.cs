using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace WordAddIn1
{
    public static class TableFormatXmlBuilder
    {
        public static string BuildGeneralOnly(TableExtractDto dto)
        {
            dto = dto ?? new TableExtractDto();
            var fontStats = dto.FontStats ?? new TableFontStatsResult();

            var generalChildren = new List<object>
            {
                new XElement("TableFontColor", fontStats.TableFontColor ?? "000000"),
            };

            if (!string.IsNullOrWhiteSpace(fontStats.TableFontName))
            {
                generalChildren.Add(new XElement("TableFontName", fontStats.TableFontName.Trim()));
            }

            if (fontStats.TableFontSize > 0)
            {
                generalChildren.Add(new XElement("TableFontSize", fontStats.TableFontSize));
            }

            string tableStyle = dto.StyleConfig?.TableStyle ?? dto.Style;
            if (!string.IsNullOrWhiteSpace(tableStyle))
            {
                int headerRows = dto.StyleConfig?.HeaderRows ?? 1;
                generalChildren.Add(
                    new XElement(
                        "Style",
                        new XElement("TableStyle", tableStyle.Trim()),
                        new XElement("HeaderRows", headerRows > 0 ? headerRows : 1)));
            }

            XDocument xmlDoc = new XDocument(
                new XElement("TableConfig", new XElement("General", generalChildren)));

            return xmlDoc.ToString();
        }

        public static string BuildDataOnly(TableExtractDto dto)
        {
            dto = dto ?? new TableExtractDto();

            var propertiesChildren = new List<object>
            {
                new XElement("Rows", dto.Rows),
                new XElement("Cols", dto.Cols),
            };

            if (dto.ColWidths != null && dto.ColWidths.Count > 0)
            {
                propertiesChildren.Insert(
                    0,
                    new XElement(
                        "ColWidths",
                        dto.ColWidths.Select(w => new XElement("float", w))));
            }

            if (dto.Merge != null && dto.Merge.Count > 0)
            {
                propertiesChildren.Add(
                    new XElement(
                        "Merge",
                        dto.Merge.Select(m => new XElement(
                            "MergeItem",
                            m.Select(v => new XElement("int", v))))));
            }

            var tableChildren = new List<object>
            {
                new XElement("Properties", propertiesChildren),
            };

            if (dto.Data != null && dto.Data.Count > 0)
            {
                tableChildren.Add(
                    new XElement(
                        "Data",
                        dto.Data.Select(row => new XElement(
                            "Row",
                            (row ?? new List<string>()).Select(BuildCellElement)))));
            }

            XDocument xmlDoc = new XDocument(
                new XElement("TableConfig", new XElement("Table", tableChildren)));

            return xmlDoc.ToString();
        }

        private static XElement BuildCellElement(string cellContent)
        {
            if (string.IsNullOrEmpty(cellContent))
            {
                return new XElement("Cell");
            }

            return new XElement("Cell", cellContent);
        }
    }
}
