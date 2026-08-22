using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordAddIn1
{
    public static class ChartXmlParser
    {
        /// <summary>
        /// CSV解析结果
        /// </summary>
        private sealed class CsvParseResult
        {
            public List<string[]> Rows { get; set; }
            public string Error { get; set; }
        }

        /// <summary>
        /// 解析XML文件
        /// </summary>
        public static async Task<ChartXmlParseResult> ParseXmlFile(string filePath)
        {
            try
            {
                string xmlContent = File.ReadAllText(filePath);
                return await ParseXmlString(xmlContent);
            }
            catch (Exception ex)
            {
                return new ChartXmlParseResult { Config = null, Error = $"读取XML文件失败: {ex.Message}" };
            }
        }

        /// <summary>
        /// 解析 ChartConfig XML 字符串（create / apply 共用）
        /// </summary>
        public static async Task<ChartXmlParseResult> ParseXmlString(string xmlContent)
        {
            try
            {
                string parsedDataSource = null;
                string parsedCsvFileName = null;

                XDocument doc = XDocument.Parse(xmlContent);
                XElement root = doc.Root;

                if (root == null || root.Name != "ChartConfig")
                {
                    return new ChartXmlParseResult { Config = null, Error = "XML根元素必须是ChartConfig" };
                }

                var config = new ChartConfig();

                // 解析Chart
                var chartElement = root.Element("Chart");
                if (chartElement != null)
                {
                    config.Chart = new ChartInfo
                    {
                        Type = chartElement.Element("Type")?.Value ?? "",
                        Title = chartElement.Element("Title")?.Value ?? "",
                        Theme = chartElement.Element("Theme")?.Value ?? "",
                        Width = int.TryParse(chartElement.Element("Width")?.Value, out int width) ? width : 0,
                        Height = int.TryParse(chartElement.Element("Height")?.Value, out int height) ? height : 0,
                        ShowDataLabels = bool.TryParse(chartElement.Element("ShowDataLabels")?.Value, out bool showDataLabels) && showDataLabels,
                        LegendPosition = chartElement.Element("LegendPosition")?.Value ?? "",
                        PlotColor = chartElement.Element("PlotColor")?.Value ?? "",
                        TitleFontSize = int.TryParse(chartElement.Element("TitleFontSize")?.Value, out int titleFontSize) ? titleFontSize : 0,
                        TitleFontBold = bool.TryParse(chartElement.Element("TitleFontBold")?.Value, out bool titleFontBold) && titleFontBold,
                        TitleFontColor = chartElement.Element("TitleFontColor")?.Value ?? "",
                        Gridlines = chartElement.Element("Gridlines")?.Value ?? "",
                        GapWidth = int.TryParse(chartElement.Element("GapWidth")?.Value, out int gapWidth) ? gapWidth : 0,
                        DataMarkers = bool.TryParse(chartElement.Element("DataMarkers")?.Value, out bool dataMarkers) && dataMarkers,
                        MarkerSize = int.TryParse(chartElement.Element("MarkerSize")?.Value, out int markerSize) ? markerSize : 0,
                        LineWeight = double.TryParse(chartElement.Element("LineWeight")?.Value, out double lineWeight) ? lineWeight : 0,
                        YAxisMin = double.TryParse(chartElement.Element("YAxisMin")?.Value, out double yAxisMin) ? yAxisMin : 0,
                        YAxisMax = double.TryParse(chartElement.Element("YAxisMax")?.Value, out double yAxisMax) ? yAxisMax : 0,
                        YAxisMajorUnit = double.TryParse(chartElement.Element("YAxisMajorUnit")?.Value, out double yAxisMajorUnit) ? yAxisMajorUnit : 0,
                        DataLabelType = chartElement.Element("DataLabelType")?.Value ?? "",
                        Explosion = int.TryParse(chartElement.Element("Explosion")?.Value, out int explosion) ? explosion : 0,
                        FillMissingValues = chartElement.Element("FillMissingValues") == null ? true : (bool.TryParse(chartElement.Element("FillMissingValues")?.Value, out bool fillMissingValues) && fillMissingValues) // 默认true，使用平滑处理
                    };

                    // 解析DataTable
                    var dataTableElement = chartElement.Element("DataTable");
                    if (dataTableElement == null)
                    {
                        return new ChartXmlParseResult { Config = config, Error = null };
                    }
                    
                    config.Chart.SeriesCollection = new List<SeriesInfo>();
                    
                    // 解析Header中的Column定义
                    var headerElement = dataTableElement.Element("Header");
                    if (headerElement == null)
                    {
                        return new ChartXmlParseResult { Config = null, Error = "DataTable中缺少Header元素" };
                    }
                    
                    var columns = headerElement.Elements("Column").ToList();
                    if (columns.Count == 0)
                    {
                        return new ChartXmlParseResult { Config = null, Error = "Header中至少需要一个Column元素" };
                    }
                    
                    // 第一个Column是category类型（X轴），其他是value类型（Y轴数据系列）
                    var valueColumns = columns.Skip(1).ToList(); // 跳过第一个category列
                    if (valueColumns.Count == 0)
                    {
                        return new ChartXmlParseResult { Config = null, Error = "Header中至少需要一个value类型的Column元素" };
                    }
                    
                    // 解析数据：优先使用DataFile（如果同时存在DataFile和Rows，优先使用DataFile），如果没有DataFile则使用Rows
                    List<string[]> dataRows = null;
                    var dataFileElement = dataTableElement.Element("DataFile");
                    var rowsElement = dataTableElement.Element("Rows");
                    
                    if (dataFileElement != null)
                    {
                        // 优先从CSV文件读取数据
                        string fileName = dataFileElement.Attribute("fileName")?.Value;
                        if (string.IsNullOrEmpty(fileName))
                        {
                            return new ChartXmlParseResult { Config = null, Error = "DataFile元素缺少fileName属性" };
                        }

                        parsedCsvFileName = fileName;

                        if (!FilePathResolver.TryResolve(fileName, out ResolvedFilePath csvResolved, out string csvError))
                        {
                            return new ChartXmlParseResult { Config = null, Error = csvError };
                        }

                        var csvRead = await FilePathResolver.ReadBytesAsync(csvResolved).ConfigureAwait(false);
                        if (!csvRead.Success || string.IsNullOrEmpty(csvResolved.LocalPath) || !File.Exists(csvResolved.LocalPath))
                        {
                            return new ChartXmlParseResult { Config = null, Error = csvRead.Error ?? csvError ?? "CSV 文件不可用" };
                        }

                        string csvFilePath = csvResolved.LocalPath;
                        
                        // 读取CSV文件
                        bool hasHeader = bool.TryParse(dataFileElement.Attribute("hasHeader")?.Value, out bool header) && header;
                        string delimiter = dataFileElement.Attribute("delimiter")?.Value ?? ",";
                        string encodingStr = dataFileElement.Attribute("encoding")?.Value ?? "UTF-8";
                        
                        var parseCsvResult = ParseCsvFile(csvFilePath, delimiter, encodingStr, hasHeader);
                        if (parseCsvResult.Error != null)
                        {
                            return new ChartXmlParseResult { Config = null, Error = $"读取CSV文件失败: {parseCsvResult.Error}" };
                        }
                        
                        dataRows = parseCsvResult.Rows;
                        if (dataRows == null || dataRows.Count == 0)
                        {
                            return new ChartXmlParseResult { Config = null, Error = "CSV文件中没有数据行" };
                        }
                    }
                    else if (rowsElement != null)
                    {
                        // 从Rows元素读取数据
                        var rows = rowsElement.Elements("Row").ToList();
                        if (rows.Count == 0)
                        {
                            return new ChartXmlParseResult { Config = null, Error = "Rows中至少需要一个Row元素" };
                        }
                        
                        dataRows = new List<string[]>();
                        foreach (var rowElement in rows)
                        {
                            string rowData = rowElement.Value ?? "";
                            string[] values = rowData.Split(',');
                            dataRows.Add(values);
                        }
                    }
                    else
                    {
                        return new ChartXmlParseResult { Config = null, Error = "DataTable中必须包含DataFile或Rows元素" };
                    }
                    
                    // 为每个value列创建一个Series
                    foreach (var columnElement in valueColumns)
                    {
                        var series = new SeriesInfo
                        {
                            Name = columnElement.Value ?? "", // Column的文本内容作为系列名称
                            AxisY = columnElement.Attribute("axisY")?.Value ?? "primary",
                            Color = columnElement.Attribute("color")?.Value ?? "",
                            ChartType = columnElement.Attribute("chartType")?.Value ?? "",
                            Points = new List<PointInfo>()
                        };
                        
                        // 解析ShowDataLabels属性（可选）
                        string showDataLabelsStr = columnElement.Attribute("showDataLabels")?.Value;
                        if (!string.IsNullOrEmpty(showDataLabelsStr) && bool.TryParse(showDataLabelsStr, out bool seriesShowDataLabels))
                        {
                            series.ShowDataLabels = seriesShowDataLabels;
                        }
                        
                        // 从数据行中提取该列的数据
                        int columnIndex = columns.IndexOf(columnElement); // 列索引（从0开始，包含category列）
                        
                        // 先收集所有数据点（原始值，无效值标记为NaN）
                        var points = new List<PointInfo>();
                        foreach (var rowValues in dataRows)
                        {
                            if (rowValues.Length > columnIndex)
                            {
                                var point = new PointInfo
                                {
                                    X = rowValues[0].Trim() // 第一列是X轴标签
                                };
                                
                                // 解析Y值（当前列的值）
                                string yValue = rowValues[columnIndex]?.Trim() ?? "";
                                
                                // 尝试解析为数字
                                if (!string.IsNullOrWhiteSpace(yValue) && double.TryParse(yValue, out double y))
                                {
                                    point.Y = y; // 有效值
                                }
                                else
                                {
                                    point.Y = double.NaN; // 使用NaN标记无效值（空值或非数字）
                                }
                                
                                points.Add(point);
                            }
                        }
                        
                        // 根据FillMissingValues配置决定处理方式
                        if (config.Chart.FillMissingValues)
                        {
                            // 使用平滑处理：前向填充 + 后向填充
                            
                            // 前向填充：用前一个有效值填充无效值
                            double? lastValidValue = null;
                            for (int i = 0; i < points.Count; i++)
                            {
                                if (double.IsNaN(points[i].Y))
                                {
                                    // 无效值，使用前一个有效值
                                    if (lastValidValue.HasValue)
                                    {
                                        points[i].Y = lastValidValue.Value;
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 数据点 {points[i].X} 的Y值为空或非数字，使用前向填充值: {points[i].Y}");
                                    }
                                }
                                else
                                {
                                    // 有效值，更新最后一个有效值
                                    lastValidValue = points[i].Y;
                                }
                            }
                            
                            // 后向填充：如果开头还有无效值，使用第一个有效值填充
                            if (points.Count > 0 && double.IsNaN(points[0].Y))
                            {
                                // 从后往前找到第一个有效值
                                double? firstValidValue = null;
                                for (int i = points.Count - 1; i >= 0; i--)
                                {
                                    if (!double.IsNaN(points[i].Y))
                                    {
                                        firstValidValue = points[i].Y;
                                        break;
                                    }
                                }
                                
                                if (firstValidValue.HasValue)
                                {
                                    // 填充开头的无效值
                                    for (int i = 0; i < points.Count && double.IsNaN(points[i].Y); i++)
                                    {
                                        points[i].Y = firstValidValue.Value;
                                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 数据点 {points[i].X} 使用后向填充值: {points[i].Y}");
                                    }
                                }
                                else
                                {
                                    // 如果整个系列都没有有效值，全部设为0
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 警告：系列 {series.Name} 没有有效的数值数据，所有值设为0");
                                    for (int i = 0; i < points.Count; i++)
                                    {
                                        points[i].Y = 0;
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 不使用平滑处理：直接将无效值设为0
                            for (int i = 0; i < points.Count; i++)
                            {
                                if (double.IsNaN(points[i].Y))
                                {
                                    points[i].Y = 0;
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 数据点 {points[i].X} 的Y值为空或非数字，设为0（未启用平滑处理）");
                                }
                            }
                        }
                        
                        // 将所有处理后的点添加到系列中
                        series.Points.AddRange(points);
                        
                        config.Chart.SeriesCollection.Add(series);
                    }

                    parsedDataSource = dataFileElement != null ? "datafile" : "rows";

                    // 解析AxisX
                    var axisXElement = chartElement.Element("AxisX");
                    if (axisXElement != null)
                    {
                        config.Chart.AxisX = new AxisXInfo
                        {
                            Title = ReadAxisField(axisXElement, "Title"),
                            Type = ReadAxisField(axisXElement, "Type", "category"),
                            Format = ReadAxisField(axisXElement, "Format"),
                            TickLabelCount = ParseOptionalInt(ReadAxisField(axisXElement, "TickLabelCount")),
                            TickLabelSpacing = ParseOptionalInt(ReadAxisField(axisXElement, "TickLabelSpacing")),
                            BetweenCategories = ParseOptionalBool(ReadAxisField(axisXElement, "BetweenCategories"))
                        };
                    }

                    // 解析AxisY
                    var axisYElement = chartElement.Element("AxisY");
                    if (axisYElement != null)
                    {
                        config.Chart.AxisY = new AxisYInfo
                        {
                            PrimaryTitle = ReadAxisField(axisYElement, "PrimaryTitle"),
                            SecondaryTitle = ReadAxisField(axisYElement, "SecondaryTitle")
                        };
                    }
                }

                return new ChartXmlParseResult
                {
                    Config = config,
                    Error = null,
                    DataSource = parsedDataSource,
                    CsvFileName = parsedCsvFileName
                };
            }
            catch (Exception ex)
            {
                string errorMsg = $"XML解析失败: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细信息: {ex.InnerException.Message}";
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] XML解析异常: {errorMsg}");
                return new ChartXmlParseResult { Config = null, Error = errorMsg };
            }
        }

        /// <summary>
        /// 解析CSV文件
        /// </summary>
        /// <param name="filePath">CSV文件路径</param>
        /// <param name="delimiter">分隔符</param>
        /// <param name="encodingStr">编码名称（如"UTF-8"）</param>
        /// <param name="hasHeader">是否包含表头</param>
        /// <returns>解析结果</returns>
        private static CsvParseResult ParseCsvFile(string filePath, string delimiter, string encodingStr, bool hasHeader)
        {
            try
            {
                // 确定编码
                Encoding encoding;
                try
                {
                    encoding = Encoding.GetEncoding(encodingStr);
                }
                catch
                {
                    encoding = Encoding.UTF8; // 默认使用UTF-8
                }

                // 读取文件所有行
                string[] lines = File.ReadAllLines(filePath, encoding);
                
                if (lines.Length == 0)
                {
                    return new CsvParseResult { Rows = new List<string[]>(), Error = null };
                }

                // 如果包含表头，跳过第一行
                int startIndex = hasHeader ? 1 : 0;
                
                var rows = new List<string[]>();
                for (int i = startIndex; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue; // 跳过空行
                    }
                    
                    // 使用指定的分隔符分割
                    string[] values = line.Split(new string[] { delimiter }, StringSplitOptions.None);
                    rows.Add(values);
                }

                return new CsvParseResult { Rows = rows, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"解析CSV文件失败: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细信息: {ex.InnerException.Message}";
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CSV解析异常: {errorMsg}");
                return new CsvParseResult { Rows = null, Error = errorMsg };
            }
        }

        private static string ReadAxisField(XElement axisElement, string name, string defaultValue = "")
        {
            if (axisElement == null)
            {
                return defaultValue;
            }

            string fromAttribute = axisElement.Attribute(name)?.Value;
            if (!string.IsNullOrEmpty(fromAttribute))
            {
                return fromAttribute;
            }

            string fromElement = axisElement.Element(name)?.Value;
            return !string.IsNullOrEmpty(fromElement) ? fromElement : defaultValue;
        }

        private static int? ParseOptionalInt(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return int.TryParse(value, out int parsed) ? (int?)parsed : null;
        }

        private static bool? ParseOptionalBool(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            if (bool.TryParse(value, out bool parsed))
            {
                return parsed;
            }

            if (value == "1")
            {
                return true;
            }

            if (value == "0")
            {
                return false;
            }

            return null;
        }
    }
}
