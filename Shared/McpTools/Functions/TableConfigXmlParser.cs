using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public static class TableConfigXmlParser
    {
        public class ParseResult
        {
            public TableApplyConfig Config { get; set; }
            public string Error { get; set; }
        }

        public class ParseDataResult
        {
            public TableDataApplyConfig Config { get; set; }
            public string Error { get; set; }
        }

        public static ParseResult Parse(string xmlContent)
        {
            try
            {
                XDocument doc = XDocument.Parse(xmlContent);
                XElement root = doc.Root;
                if (root == null || root.Name != "TableConfig")
                {
                    return new ParseResult { Config = null, Error = "XML根元素必须是TableConfig" };
                }

                var config = new TableApplyConfig
                {
                    General = new TableGeneralFormat(),
                    Properties = new TablePropertiesFormat(),
                    ColWidths = new List<float>(),
                    Merge = new List<List<int>>(),
                };

                XElement generalElement = root.Element("General");
                if (generalElement?.Element("ColWidths") != null)
                {
                    return new ParseResult
                    {
                        Config = null,
                        Error = "TableConfig 布局无效：ColWidths 须位于 Table/Properties 下，请重新提取表格格式",
                    };
                }

                XElement tableElement = root.Element("Table");
                if (tableElement?.Element("ColWidths") != null)
                {
                    return new ParseResult
                    {
                        Config = null,
                        Error = "TableConfig 布局无效：ColWidths 须位于 Table/Properties 下，请重新提取表格格式",
                    };
                }

                if (generalElement != null)
                {
                    config.General.TableFontColor = generalElement.Element("TableFontColor")?.Value?.Trim();
                    XElement nameEl = generalElement.Element("TableFontName");
                    if (nameEl != null && !string.IsNullOrWhiteSpace(nameEl.Value))
                    {
                        config.General.TableFontName = nameEl.Value.Trim();
                    }

                    XElement sizeEl = generalElement.Element("TableFontSize");
                    if (sizeEl != null &&
                        float.TryParse(sizeEl.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out float fs) &&
                        fs > 0)
                    {
                        config.General.TableFontSize = fs;
                    }

                    config.General.StyleConfig = ParseStyleElement(generalElement.Element("Style"));
                }

                if (tableElement != null)
                {
                    XElement propertiesElement = tableElement.Element("Properties");
                    if (propertiesElement?.Element("Style") != null)
                    {
                        return new ParseResult
                        {
                            Config = null,
                            Error = "TableConfig 布局无效：Style 须位于 General 下（Properties 内不再支持），请重新提取表格格式",
                        };
                    }

                    if (propertiesElement != null)
                    {
                        List<float> colWidths = ParseColWidthsElement(propertiesElement.Element("ColWidths"));
                        if (colWidths.Count > 0)
                        {
                            config.ColWidths = colWidths;
                        }

                        if (int.TryParse(propertiesElement.Element("Rows")?.Value, out int rows))
                        {
                            config.Properties.Rows = rows;
                        }
                        if (int.TryParse(propertiesElement.Element("Cols")?.Value, out int cols))
                        {
                            config.Properties.Cols = cols;
                        }

                        XElement mergeElement = propertiesElement.Element("Merge");
                        if (mergeElement != null)
                        {
                            foreach (XElement mergeItemElement in mergeElement.Elements("MergeItem"))
                            {
                                var mergeItem = new List<int>();
                                foreach (XElement intElement in mergeItemElement.Elements("int"))
                                {
                                    if (int.TryParse(intElement.Value, out int value))
                                    {
                                        mergeItem.Add(value);
                                    }
                                }
                                if (mergeItem.Count == 4)
                                {
                                    config.Merge.Add(mergeItem);
                                }
                            }
                        }
                    }
                }

                return new ParseResult { Config = config, Error = null };
            }
            catch (Exception ex)
            {
                return new ParseResult { Config = null, Error = $"XML解析失败: {ex.Message}" };
            }
        }

        public static ParseDataResult ParseDataXml(string xmlContent)
        {
            try
            {
                XDocument doc = XDocument.Parse(xmlContent);
                XElement root = doc.Root;
                if (root == null)
                {
                    return new ParseDataResult { Config = null, Error = "XML根元素无效" };
                }

                if (root.Name != "TableConfig" && root.Name != "Table")
                {
                    return new ParseDataResult { Config = null, Error = "XML根元素必须是TableConfig或Table" };
                }

                XElement tableElement = root.Name == "Table" ? root : root.Element("Table");
                if (tableElement == null)
                {
                    return new ParseDataResult { Config = null, Error = "XML必须含 Table 元素" };
                }

                XElement dataElement = tableElement.Element("Data");
                if (dataElement == null)
                {
                    return new ParseDataResult { Config = null, Error = "XML必须含 Table/Data" };
                }

                var config = new TableDataApplyConfig
                {
                    Properties = new TablePropertiesFormat(),
                    Merge = new List<List<int>>(),
                    Data = new List<List<string>>(),
                };

                XElement propertiesElement = tableElement.Element("Properties");
                if (propertiesElement != null)
                {
                    List<float> colWidths = ParseColWidthsElement(propertiesElement.Element("ColWidths"));
                    if (colWidths.Count > 0)
                    {
                        config.ColWidths = colWidths;
                    }

                    if (int.TryParse(propertiesElement.Element("Rows")?.Value, out int rows))
                    {
                        config.Properties.Rows = rows;
                    }

                    if (int.TryParse(propertiesElement.Element("Cols")?.Value, out int cols))
                    {
                        config.Properties.Cols = cols;
                    }

                    XElement mergeElement = propertiesElement.Element("Merge");
                    if (mergeElement != null)
                    {
                        foreach (XElement mergeItemElement in mergeElement.Elements("MergeItem"))
                        {
                            var mergeItem = new List<int>();
                            foreach (XElement intElement in mergeItemElement.Elements("int"))
                            {
                                if (int.TryParse(intElement.Value, out int value))
                                {
                                    mergeItem.Add(value);
                                }
                            }

                            if (mergeItem.Count == 4)
                            {
                                config.Merge.Add(mergeItem);
                            }
                        }
                    }
                }

                foreach (XElement rowElement in dataElement.Elements("Row"))
                {
                    var row = new List<string>();
                    foreach (XElement cellElement in rowElement.Elements("Cell"))
                    {
                        row.Add(ParseCellElementValue(cellElement));
                    }

                    config.Data.Add(row);
                }

                if (!config.Properties.Rows.HasValue && config.Data.Count > 0)
                {
                    config.Properties.Rows = config.Data.Count;
                }

                if (!config.Properties.Cols.HasValue && config.Data.Count > 0)
                {
                    config.Properties.Cols = config.Data.Max(r => r?.Count ?? 0);
                }

                return new ParseDataResult { Config = config, Error = null };
            }
            catch (Exception ex)
            {
                return new ParseDataResult { Config = null, Error = $"XML解析失败: {ex.Message}" };
            }
        }

        private static string ParseCellElementValue(XElement cellElement)
        {
            if (cellElement == null)
            {
                return "";
            }

            if (cellElement.HasElements)
            {
                return string.Join("", cellElement.Nodes().Select(n => n.ToString()));
            }

            return cellElement.Value ?? "";
        }

        public static TableApplyConfig MergeWithExisting(TableApplyConfig existing, TableApplyConfig target)
        {
            if (existing == null)
            {
                return target;
            }
            if (target == null)
            {
                return existing;
            }

            var merged = new TableApplyConfig
            {
                General = new TableGeneralFormat(),
                Properties = new TablePropertiesFormat(),
                Merge = target.Merge != null && target.Merge.Count > 0 ? target.Merge : existing.Merge,
            };

            merged.ColWidths = target.ColWidths != null && target.ColWidths.Count > 0
                ? target.ColWidths
                : existing.ColWidths;
            merged.General.TableFontName = target.General?.TableFontName ?? existing.General?.TableFontName;
            merged.General.TableFontSize = target.General?.TableFontSize ?? existing.General?.TableFontSize;
            merged.General.TableFontColor = target.General?.TableFontColor ?? existing.General?.TableFontColor;
            merged.General.StyleConfig = target.General?.StyleConfig ?? existing.General?.StyleConfig;

            merged.Properties.Rows = target.Properties?.Rows ?? existing.Properties?.Rows;
            merged.Properties.Cols = target.Properties?.Cols ?? existing.Properties?.Cols;

            return merged;
        }

        private static TableStyleConfig ParseStyleElement(XElement styleElement)
        {
            if (styleElement == null)
            {
                return null;
            }

            string tableStyle = styleElement.Element("TableStyle")?.Value?.Trim() ?? "";
            if (string.IsNullOrEmpty(tableStyle))
            {
                return null;
            }

            int headerRows = int.TryParse(styleElement.Element("HeaderRows")?.Value, out int hr) ? hr : 1;
            return new TableStyleConfig
            {
                TableStyle = tableStyle,
                HeaderRows = Math.Max(1, headerRows),
            };
        }

        private static List<float> ParseColWidthsElement(XElement colWidthsElement)
        {
            var widths = new List<float>();
            if (colWidthsElement == null)
            {
                return widths;
            }

            foreach (XElement widthElement in colWidthsElement.Elements("float"))
            {
                if (float.TryParse(widthElement.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out float width))
                {
                    widths.Add(width);
                }
            }

            return widths;
        }
    }
}
