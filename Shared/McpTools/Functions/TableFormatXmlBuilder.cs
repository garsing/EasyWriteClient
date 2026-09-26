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
                tableChildren.Add(BuildOmitDataElement(dto));
            }

            XDocument xmlDoc = new XDocument(
                new XElement("TableConfig", new XElement("Table", tableChildren)));

            return xmlDoc.ToString();
        }

        public static string BuildGeneralWithOmitData(TableExtractDto dto)
        {
            dto = dto ?? new TableExtractDto();
            XDocument generalDoc = XDocument.Parse(BuildGeneralOnly(dto));
            XDocument dataDoc = XDocument.Parse(BuildDataOnly(dto));
            XElement root = generalDoc.Root;
            XElement table = dataDoc.Root?.Element("Table");
            if (root != null && table != null)
            {
                root.Add(table);
            }

            return generalDoc.ToString();
        }

        private static XElement BuildOmitDataElement(TableExtractDto dto)
        {
            int cols = dto.Cols;
            if (cols <= 0 && dto.Data != null)
            {
                foreach (List<string> row in dto.Data)
                {
                    if (row != null && row.Count > cols)
                    {
                        cols = row.Count;
                    }
                }
            }

            var rowEls = new List<XElement>();
            for (int r = 0; r < dto.Data.Count; r++)
            {
                List<string> row = dto.Data[r] ?? new List<string>();
                var cells = new List<XElement>();
                int colCount = cols > 0 ? cols : row.Count;
                for (int c = 1; c <= colCount; c++)
                {
                    if (TableMergeHelper.IsMergeContinuation(r + 1, c, dto.Merge))
                    {
                        continue;
                    }

                    string text = c - 1 < row.Count ? row[c - 1] : "";
                    XElement cellEl = BuildCellElement(text);
                    int rowSpan;
                    int colSpan;
                    TryGetMergeSpan(r + 1, c, dto.Merge, out rowSpan, out colSpan);
                    if (rowSpan > 1)
                    {
                        cellEl.SetAttributeValue("rowspan", rowSpan);
                    }

                    if (colSpan > 1)
                    {
                        cellEl.SetAttributeValue("colspan", colSpan);
                    }

                    cells.Add(cellEl);
                }

                rowEls.Add(new XElement("Row", cells.Cast<object>().ToArray()));
            }

            return new XElement("Data", rowEls.Cast<object>().ToArray());
        }

        private static void TryGetMergeSpan(
            int row,
            int col,
            List<List<int>> merge,
            out int rowSpan,
            out int colSpan)
        {
            rowSpan = 1;
            colSpan = 1;
            if (merge == null)
            {
                return;
            }

            foreach (List<int> item in merge)
            {
                if (item == null || item.Count < 4)
                {
                    continue;
                }

                if (item[0] == row && item[1] == col)
                {
                    rowSpan = item[2] - item[0] + 1;
                    colSpan = item[3] - item[1] + 1;
                    if (rowSpan < 1)
                    {
                        rowSpan = 1;
                    }

                    if (colSpan < 1)
                    {
                        colSpan = 1;
                    }

                    return;
                }
            }
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
