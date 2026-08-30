using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace WordAddIn1
{
    /// <summary>
    /// Word文档读取器
    /// 用于读取Word文档内容，包括段落、表格等
    /// </summary>
    public static class WordReader
    {
        // 图片计数器：key = 文档路径hash, value = 图片计数
        private static Dictionary<string, int> _imageCounters = new Dictionary<string, int>();
        private static int _placeholderImageCounter;

        // 单次 ReadWord 内的完整抽图上下文（按 XML 顺序逐 drawing 落盘）
        private static bool _readUseFullImages;
        private static string _readDocPath;
        private static string _readImageSaveDir;
        private static string _readFileHash;
        private static ZipArchive _readZip;
        private static Dictionary<string, string> _readRelsDict;

        private static readonly XNamespace NS_CHART =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private sealed class ChartOrderedEntry
        {
            public string TaggedXml { get; set; }
        }

        private static List<ChartOrderedEntry> _chartOrderedEntries;
        private static int _chartConsumeIndex;

        private static HashSet<string> _readOccupiedTableIds;
        private static List<string> _readWordAssignedIds;
        private static int _readTableEmitIndex;

        private static void LogExtractVerbose(string message)
        {
            EasyWriteDiagnostics.LogDocumentExtractVerbose(message);
        }

        private const string DebugParagraphKeyword = "图6展示";

        private static bool IsRevisionInsertion(XElement el, XNamespace w)
        {
            return el != null && (el.Name == w + "ins" || el.Name.LocalName == "ins"
                || el.Name == w + "moveTo" || el.Name.LocalName == "moveTo");
        }

        private static bool IsRevisionDeletion(XElement el, XNamespace w)
        {
            return el != null && (el.Name == w + "del" || el.Name.LocalName == "del"
                || el.Name == w + "delText" || el.Name.LocalName == "delText"
                || el.Name == w + "moveFrom" || el.Name.LocalName == "moveFrom");
        }

        /// <summary>
        /// 表格合并信息跟踪上下文
        /// </summary>
        private class TableMergeContext
        {
            /// <summary>
            /// 纵向合并映射：key = 逻辑列索引, value = (起始行索引, gridSpan值)
            /// </summary>
            public Dictionary<int, (int startRow, int gridSpan)> VerticalMerges { get; set; }
            
            /// <summary>
            /// 当前行索引（从0开始）
            /// </summary>
            public int CurrentRowIndex { get; set; }
            
            /// <summary>
            /// 当前逻辑列索引（考虑gridSpan后的实际列位置）
            /// </summary>
            public int CurrentLogicalColumnIndex { get; set; }
            
            public TableMergeContext()
            {
                VerticalMerges = new Dictionary<int, (int, int)>();
                CurrentRowIndex = -1;
                CurrentLogicalColumnIndex = 0;
            }
        }

        /// <summary>
        /// 读取Word文档内容
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="docPath">文档路径（可选，用于图片提取）</param>
        /// <returns>文档内容文本</returns>
        public static string ReadWord(Word.Document document, string docPath = null)
        {
            return ReadWord(document, docPath, extractImages: true);
        }

        /// <param name="extractImages">为 false 时不抽图、不扫图表（process_actions 轻量路径）。</param>
        public static string ReadWord(Word.Document document, string docPath, bool extractImages)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            try
            {
                // 每次运行都清空表格/图表/图片顺序表，重新生成当前文档的顺序
                DocumentState.ClearTableIdOrder();
                DocumentState.ClearChartIdOrder();
                DocumentState.ClearImageIdOrder();
                DocumentState.ClearImageIdMapping(); // 图片按实例唯一，每次读文重建
                _readOccupiedTableIds = new HashSet<string>(StringComparer.Ordinal);
                _readWordAssignedIds = TableTitleStampHelper.CollectIdsFromWord(document, _readOccupiedTableIds);
                _readTableEmitIndex = 0;
                _placeholderImageCounter = 0;
                
                // 直接从文档的 XML 提取段落和表格，避免重复
                // 获取整个文档的 XML
                string docXml = document.Content.WordOpenXML;
                if (string.IsNullOrEmpty(docXml))
                {
                    throw new Exception("无法获取文档的 XML 内容");
                }

                // 解析 XML
                XDocument xmlDoc = XDocument.Parse(docXml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                // 查找文档主体（body）
                var body = xmlDoc.Descendants(w + "body").FirstOrDefault();
                if (body == null)
                {
                    throw new Exception("无法找到文档主体");
                }

                string fileHash = null;
                string imageSaveDir = null;
                string pathForImages = null;

                // 图表走 COM 预提取，OOXML 遍历时按 chart drawing 顺序消费
                _chartOrderedEntries = ExtractAllChartsOrdered(document);
                _chartConsumeIndex = 0;

                if (extractImages)
                {
                    pathForImages = string.IsNullOrEmpty(docPath) ? document.FullName : docPath;
                    string docUuid = DocumentState.CurrDocUuid;

                    EasyWriteDiagnostics.LogDocumentExtractVerbose($"[WordReader] 文档路径: {pathForImages}");
                    EasyWriteDiagnostics.LogDocumentExtractVerbose($"[WordReader] 文档UUID: {docUuid}");

                    fileHash = CalculateFileHash(pathForImages);
                    EasyWriteDiagnostics.LogDocumentExtractVerbose($"[WordReader] 文件hash: {fileHash}");

                    imageSaveDir = GetImageSaveDirectory(pathForImages, fileHash);
                    EasyWriteDiagnostics.LogDocumentExtractVerbose($"[WordReader] 图片保存目录: {imageSaveDir}");
                }
                else
                {
                    EasyWriteDiagnostics.LogDocumentExtractVerbose(
                        "[WordReader] 轻量模式：不 SaveAs/不解 zip，插入图片占位标签以保持 seg 结构");
                }

                BeginReadImageContext(extractImages, pathForImages, imageSaveDir, fileHash);
                string result;
                try
                {
                    // 递归处理 body 的所有子元素，一步到位提取所有内容
                    result = ProcessElementRecursive(body, w, document, pathForImages, imageSaveDir, fileHash);
                }
                finally
                {
                    EndReadImageContext();
                }

                BuildTableIdOrderFromText(result);
                result = ApplyTableStampsAfterRead(document, result);
                BuildChartIdOrderFromText(result);
                BuildImageIdOrderFromText(result);
                
                return result;
            }
            catch (Exception ex)
            {
                throw new Exception($"Word文本提取失败: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 本份 ReadText 第 N 张表的号来自 Word 第 N 张 Title。COM 少收时才退回该 w:tbl 自己的 caption。
        /// </summary>
        private static string TakeWordSourcedTableId(string tableContent, XElement tbl, XNamespace w)
        {
            string tableId = null;
            int i = _readTableEmitIndex;
            if (_readWordAssignedIds != null && i < _readWordAssignedIds.Count)
            {
                tableId = _readWordAssignedIds[i];
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[TableStamp] ReadText 第 {i} 张超出 Word 表数 {_readWordAssignedIds?.Count ?? 0}，退回 XML caption");
                tableId = TableTitleStampHelper.DecideTableId(ReadTblCaption(tbl, w), _readOccupiedTableIds);
            }

            if (string.IsNullOrEmpty(tableId))
            {
                tableId = DocumentState.FindOrCreateTableId(tableContent);
                if (_readOccupiedTableIds != null)
                {
                    _readOccupiedTableIds.Add(tableId);
                }
            }

            if (_readWordAssignedIds != null && i < _readWordAssignedIds.Count)
            {
                _readWordAssignedIds[i] = tableId;
            }

            _readTableEmitIndex++;
            return tableId;
        }

        private static string ReadTblCaption(XElement tbl, XNamespace w)
        {
            if (tbl == null)
            {
                return "";
            }

            XElement tblPr = tbl.Element(w + "tblPr");
            XElement caption = tblPr?.Element(w + "tblCaption");
            if (caption == null)
            {
                return "";
            }

            XAttribute val = caption.Attribute(w + "val") ?? caption.Attribute("val");
            return val != null ? (val.Value ?? "") : "";
        }

        private static string ApplyTableStampsAfterRead(Word.Document document, string result)
        {
            List<string> ids = DocumentState.GetTableIdOrder();
            if (ids == null || ids.Count == 0)
            {
                return result;
            }

            var before = new List<string>(ids);
            TableTitleStampHelper.ApplyAssignedIds(document, ids, tableContents: null);
            bool changed = false;
            for (int i = 0; i < ids.Count; i++)
            {
                if (i < before.Count && ids[i] != before[i])
                {
                    result = TableTitleStampHelper.ReplaceTableIdAtOccurrence(result, i, ids[i]);
                    changed = true;
                }
            }

            if (changed)
            {
                var map = new Dictionary<int, string>();
                for (int i = 0; i < ids.Count; i++)
                {
                    map[i + 1] = ids[i];
                }

                DocumentState.SetTableIndexToIdMap(map);
            }

            return result;
        }

        /// <summary>
        /// 从生成的文本中按照 <table id="..."> 标签的出现顺序建立表格ID顺序映射表
        /// 这样可以确保顺序和文档中实际出现的顺序一致（包括嵌套表格）
        /// </summary>
        /// <param name="text">生成的文本内容</param>
        private static void BuildTableIdOrderFromText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            // 使用正则表达式匹配所有 <table id="..."> 标签
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"<table\s+id=""([^""]+)""");
            var matches = regex.Matches(text);

            // 使用 HashSet 来跟踪已经添加的 tableId，避免重复添加
            var addedTableIds = new System.Collections.Generic.HashSet<string>();

            // 按照在文本中出现的顺序，将表格ID添加到顺序数组中
            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (match.Groups.Count >= 2)
                {
                    string tableId = match.Groups[1].Value;
                    // 如果同一个tableId还没有添加过，添加到顺序数组中（保持第一次出现的位置）
                    if (!string.IsNullOrEmpty(tableId) && !addedTableIds.Contains(tableId))
                    {
                        DocumentState.AddTableIdToOrder(tableId);
                        addedTableIds.Add(tableId);
                    }
                }
            }
        }

        private static void BuildChartIdOrderFromText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var regex = new System.Text.RegularExpressions.Regex(@"<Chart\s+id=""([^""]+)""");
            var matches = regex.Matches(text);
            var addedChartIds = new System.Collections.Generic.HashSet<string>();

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (match.Groups.Count >= 2)
                {
                    string chartId = match.Groups[1].Value;
                    if (!string.IsNullOrEmpty(chartId) && !addedChartIds.Contains(chartId))
                    {
                        DocumentState.AddChartIdToOrder(chartId);
                        addedChartIds.Add(chartId);
                    }
                }
            }
        }

        private static void BuildImageIdOrderFromText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            var regex = new System.Text.RegularExpressions.Regex(@"<image\b[^>]*\bid=""([^""]+)""");
            var matches = regex.Matches(text);
            var addedImageIds = new System.Collections.Generic.HashSet<string>();

            foreach (System.Text.RegularExpressions.Match match in matches)
            {
                if (match.Groups.Count >= 2)
                {
                    string imageId = match.Groups[1].Value;
                    if (!string.IsNullOrEmpty(imageId) && !addedImageIds.Contains(imageId))
                    {
                        DocumentState.AddImageIdToOrder(imageId);
                        addedImageIds.Add(imageId);
                    }
                }
            }
        }

        private static List<ChartOrderedEntry> ExtractAllChartsOrdered(Word.Document document)
        {
            var orderedShapes = new List<(int Start, Word.Chart Chart)>();

            try
            {
                foreach (Word.InlineShape shape in document.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapeChart)
                    {
                        try
                        {
                            orderedShapes.Add((shape.Range.Start, shape.Chart));
                        }
                        catch (Exception ex)
                        {
                            LogExtractVerbose($"[WordReader] 读取 InlineShape 图表失败: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogExtractVerbose($"[WordReader] 提取图表列表失败: {ex.Message}");
                return new List<ChartOrderedEntry>();
            }

            orderedShapes.Sort((a, b) => a.Start.CompareTo(b.Start));

            var result = new List<ChartOrderedEntry>();
            foreach (var item in orderedShapes)
            {
                string rawXml = ExtractChartSeriesCollection(item.Chart);
                if (string.IsNullOrEmpty(rawXml))
                {
                    continue;
                }

                string bodyContent = ExtractChartBodyContent(rawXml);
                string chartId = DocumentState.FindOrCreateChartId(bodyContent);
                string taggedXml = BuildChartXmlWithId(chartId, rawXml);
                result.Add(new ChartOrderedEntry { TaggedXml = taggedXml });
                LogExtractVerbose($"[ChartId] 预提取 chart_id={chartId}，Range.Start={item.Start}");
            }

            LogExtractVerbose($"[WordReader] 共预提取 {result.Count} 个图表（有序）");
            return result;
        }

        private static string ExtractChartBodyContent(string rawChartXml)
        {
            if (string.IsNullOrEmpty(rawChartXml))
            {
                return "";
            }

            int start = rawChartXml.IndexOf("<DataTable", StringComparison.OrdinalIgnoreCase);
            if (start < 0)
            {
                return rawChartXml.Trim();
            }

            int end = rawChartXml.LastIndexOf("</DataTable>", StringComparison.OrdinalIgnoreCase);
            if (end < start)
            {
                return rawChartXml.Trim();
            }

            end += "</DataTable>".Length;
            return rawChartXml.Substring(start, end - start).Trim();
        }

        private static string BuildChartXmlWithId(string chartId, string rawChartXml)
        {
            if (string.IsNullOrEmpty(rawChartXml))
            {
                return "";
            }

            if (rawChartXml.StartsWith("<Chart", StringComparison.OrdinalIgnoreCase))
            {
                int closeBracket = rawChartXml.IndexOf('>');
                if (closeBracket > 0)
                {
                    return "<Chart id=\"" + chartId + "\">" + rawChartXml.Substring(closeBracket + 1);
                }
            }

            return "<Chart id=\"" + chartId + "\">\n" + rawChartXml + "</Chart>\n";
        }

        private static string TryConsumeNextChartTaggedXml()
        {
            if (_chartOrderedEntries == null || _chartConsumeIndex >= _chartOrderedEntries.Count)
            {
                return "";
            }

            string xml = _chartOrderedEntries[_chartConsumeIndex].TaggedXml;
            _chartConsumeIndex++;
            return xml;
        }

        private static bool IsChartDrawing(XElement drawing)
        {
            if (drawing == null)
            {
                return false;
            }

            return drawing.Descendants(NS_CHART + "chart").Any()
                || drawing.Descendants(NS_CHART + "chartSpace").Any()
                || drawing.Descendants(NS_CHART + "userShapes").Any();
        }

        private static string AppendDrawingContent(XElement drawing, XNamespace w)
        {
            if (IsChartDrawing(drawing))
            {
                string chartXml = TryConsumeNextChartTaggedXml();
                if (!string.IsNullOrEmpty(chartXml))
                {
                    return chartXml + "\r";
                }

                return "";
            }

            string meta = ExtractImageMetaFromDrawing(drawing);
            string imageId = DocumentState.AllocateUniqueImageId(meta);
            return FormatImageTag(imageId, ResolveImagePathForDrawing(drawing));
        }

        /// <summary>
        /// 规范化拼串：docPrId|docPrName|cx|cy|rId（缺省空串；rId 优先 embed）。
        /// </summary>
        private static string ExtractImageMetaFromDrawing(XElement drawing)
        {
            if (drawing == null)
            {
                return "||||";
            }

            XNamespace nsWp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
            XNamespace nsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

            var docPr = drawing.Descendants(nsWp + "docPr").FirstOrDefault();
            string docPrId = docPr?.Attribute("id")?.Value ?? "";
            string docPrName = docPr?.Attribute("name")?.Value ?? "";

            var extent = drawing.Descendants(nsWp + "extent").FirstOrDefault();
            if (extent == null)
            {
                extent = drawing.Descendants(nsA + "ext").FirstOrDefault();
            }

            string cx = extent?.Attribute("cx")?.Value ?? "";
            string cy = extent?.Attribute("cy")?.Value ?? "";

            var blip = drawing.Descendants(nsA + "blip").FirstOrDefault();
            string rId = "";
            if (blip != null)
            {
                rId = blip.Attribute(nsR + "embed")?.Value
                    ?? blip.Attribute(nsR + "link")?.Value
                    ?? "";
            }

            return $"{docPrId}|{docPrName}|{cx}|{cy}|{rId}";
        }

        /// <summary>
        /// 提取图表的DataTable数据（格式化为XML）
        /// </summary>
        /// <param name="chart">Word图表对象</param>
        /// <returns>DataTable的XML字符串</returns>
        private static string ExtractChartSeriesCollection(Word.Chart chart)
        {
            try
            {
                dynamic seriesCollection = chart.SeriesCollection();
                int seriesCount = seriesCollection.Count;

                if (seriesCount == 0)
                {
                    return "<Chart>\n  <DataTable>\n  </DataTable>\n</Chart>\n";
                }

                // 收集所有Series的信息和数据
                var seriesInfoList = new List<SeriesInfo>();
                var allXValues = new List<string>(); // 收集所有唯一的X值
                var xValueSet = new HashSet<string>();

                for (int i = 1; i <= seriesCount; i++)
                {
                    try
                    {
                        dynamic series = seriesCollection(i);
                        
                        var seriesInfo = new SeriesInfo
                        {
                            Name = series.Name ?? $"系列{i}",
                            AxisY = "primary",
                            Color = "",
                            ChartType = "",
                            ShowDataLabels = "false",
                            Points = new List<PointData>()
                        };
                        
                        // 获取AxisY（primary或secondary）
                        try
                        {
                            dynamic axisGroup = series.AxisGroup;
                            if (axisGroup == Word.XlAxisGroup.xlSecondary)
                            {
                                seriesInfo.AxisY = "secondary";
                            }
                        }
                        catch { }
                        
                        // 获取颜色
                        try
                        {
                            // 尝试获取线条颜色（折线图）
                            try
                            {
                                int rgbValue = series.Format.Line.ForeColor.RGB;
                                seriesInfo.Color = ConvertRgbToHex(rgbValue);
                            }
                            catch
                            {
                                // 尝试获取填充颜色（柱状图、饼图）
                                try
                                {
                                    int rgbValue = series.Format.Fill.ForeColor.RGB;
                                    seriesInfo.Color = ConvertRgbToHex(rgbValue);
                                }
                                catch { }
                            }
                        }
                        catch { }
                        
                        // 获取ChartType（用于混合图表）
                        try
                        {
                            dynamic chartTypeValue = series.ChartType;
                            if (chartTypeValue == Excel.XlChartType.xlColumnClustered)
                            {
                                seriesInfo.ChartType = "column";
                            }
                            else if (chartTypeValue == Excel.XlChartType.xlLine)
                            {
                                seriesInfo.ChartType = "line";
                            }
                            else if (chartTypeValue == Excel.XlChartType.xl3DPie)
                            {
                                seriesInfo.ChartType = "pie3d";
                            }
                        }
                        catch { }
                        
                        // 获取ShowDataLabels
                        try
                        {
                            bool hasDataLabels = series.HasDataLabels;
                            seriesInfo.ShowDataLabels = hasDataLabels ? "true" : "false";
                        }
                        catch { }

                        // 获取X轴标签和Y轴值
                        try
                        {
                            object valuesObj = null;
                            object xValuesObj = null;
                            try
                            {
                                var seriesObj = (object)series;
                                var t = seriesObj.GetType();
                                valuesObj = t.InvokeMember("Values", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, seriesObj, null);
                                xValuesObj = t.InvokeMember("XValues", System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Instance, null, seriesObj, null);
                            }
                            catch (Exception ex)
                            {
                                LogExtractVerbose($"[WordReader] 通过反射获取 Values/XValues 失败: {ex.Message}");
                            }
                            
                            int pointCount = 0;
                            Array valuesArray = null;
                            Array xValuesArray = null;
                            
                            // 处理 values（Y值）
                            if (valuesObj != null && valuesObj is Array arr)
                            {
                                valuesArray = arr;
                                int lowerBound = valuesArray.GetLowerBound(0);
                                int upperBound = valuesArray.GetUpperBound(0);
                                pointCount = upperBound - lowerBound + 1;
                            }
                            
                            // 处理 xValues（X值）
                            if (xValuesObj != null && xValuesObj is Array xArr)
                            {
                                xValuesArray = xArr;
                            }
                            
                            // 确定数组的起始索引
                            int valuesStartIndex = 0;
                            int xValuesStartIndex = 0;
                            if (valuesArray != null)
                            {
                                valuesStartIndex = valuesArray.GetLowerBound(0);
                            }
                            if (xValuesArray != null)
                            {
                                xValuesStartIndex = xValuesArray.GetLowerBound(0);
                            }
                            
                            for (int j = 0; j < pointCount; j++)
                            {
                                string xValue = "";
                                double yValue = 0;

                                // 获取X值
                                if (xValuesArray != null)
                                {
                                    int arrayIndex = xValuesStartIndex + j;
                                    if (arrayIndex >= xValuesArray.GetLowerBound(0) && arrayIndex <= xValuesArray.GetUpperBound(0))
                                    {
                                        object xVal = xValuesArray.GetValue(arrayIndex);
                                        xValue = xVal?.ToString() ?? "";
                                        if (!string.IsNullOrEmpty(xValue) && !xValueSet.Contains(xValue))
                                        {
                                            xValueSet.Add(xValue);
                                            allXValues.Add(xValue);
                                        }
                                    }
                                }

                                // 获取Y值
                                if (valuesArray != null)
                                {
                                    int arrayIndex = valuesStartIndex + j;
                                    if (arrayIndex >= valuesArray.GetLowerBound(0) && arrayIndex <= valuesArray.GetUpperBound(0))
                                    {
                                        object val = valuesArray.GetValue(arrayIndex);
                                        if (val != null && double.TryParse(val.ToString(), out double y))
                                            yValue = y;
                                    }
                                }

                                seriesInfo.Points.Add(new PointData { X = xValue, Y = yValue });
                            }
                            
                            if (pointCount == 0)
                            {
                                LogExtractVerbose($"[WordReader] ⚠️ 系列{i}没有数据点（pointCount=0）");
                            }
                            else
                            {
                                LogExtractVerbose($"[WordReader] ✅ 系列{i}成功提取 {pointCount} 个数据点");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogExtractVerbose($"[WordReader] ❌ 提取系列{i}的数据点失败: {ex.Message}");
                        }

                        seriesInfoList.Add(seriesInfo);
                    }
                    catch (Exception ex)
                    {
                        LogExtractVerbose($"[WordReader] 提取系列{i}失败: {ex.Message}");
                    }
                }

                // 构建DataTable XML
                var dataTableXml = new StringBuilder();
                dataTableXml.Append("<Chart>\n");
                dataTableXml.Append("  <DataTable>\n");
                dataTableXml.Append("    <Header>\n");
                
                // 第一个Column是category类型（X轴）
                dataTableXml.Append("      <Column type=\"category\">类别</Column>\n");
                
                // 其他Column是value类型（Y轴数据系列）
                foreach (var seriesInfo in seriesInfoList)
                {
                    dataTableXml.Append("      <Column");
                    dataTableXml.Append($" type=\"value\"");
                    dataTableXml.Append($" axisY=\"{seriesInfo.AxisY}\"");
                    if (!string.IsNullOrEmpty(seriesInfo.Color))
                    {
                        dataTableXml.Append($" color=\"{seriesInfo.Color}\"");
                    }
                    if (!string.IsNullOrEmpty(seriesInfo.ChartType))
                    {
                        dataTableXml.Append($" chartType=\"{seriesInfo.ChartType}\"");
                    }
                    if (!string.IsNullOrEmpty(seriesInfo.ShowDataLabels))
                    {
                        dataTableXml.Append($" showDataLabels=\"{seriesInfo.ShowDataLabels}\"");
                    }
                    dataTableXml.Append($">{EscapeXmlText(seriesInfo.Name)}</Column>\n");
                }
                
                dataTableXml.Append("    </Header>\n");
                dataTableXml.Append("    <Rows>\n");
                
                // 构建Rows：每行是逗号分隔的值
                // 使用第一个Series的X值作为基准（假设所有Series的X值一致）
                if (seriesInfoList.Count > 0 && seriesInfoList[0].Points.Count > 0)
                {
                    int rowCount = seriesInfoList[0].Points.Count;
                    for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
                    {
                        var rowValues = new List<string>();
                        
                        // 第一个值是X轴标签
                        string xValue = seriesInfoList[0].Points[rowIndex].X ?? "";
                        rowValues.Add(xValue);
                        
                        // 后续值是各个Series的Y值
                        foreach (var seriesInfo in seriesInfoList)
                        {
                            if (rowIndex < seriesInfo.Points.Count)
                            {
                                double yValue = seriesInfo.Points[rowIndex].Y;
                                rowValues.Add(yValue.ToString("F2"));
                            }
                            else
                            {
                                rowValues.Add("0");
                            }
                        }
                        
                        string rowData = string.Join(",", rowValues);
                        dataTableXml.Append($"      <Row>{EscapeXmlText(rowData)}</Row>\n");
                    }
                }
                
                dataTableXml.Append("    </Rows>\n");
                dataTableXml.Append("  </DataTable>\n");
                dataTableXml.Append("</Chart>\n");

                return dataTableXml.ToString();
            }
            catch (Exception ex)
            {
                LogExtractVerbose($"[WordReader] 提取图表DataTable失败: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// Series信息辅助类
        /// </summary>
        private class SeriesInfo
        {
            public string Name { get; set; }
            public string AxisY { get; set; }
            public string Color { get; set; }
            public string ChartType { get; set; }
            public string ShowDataLabels { get; set; }
            public List<PointData> Points { get; set; }
        }

        /// <summary>
        /// 数据点辅助类
        /// </summary>
        private class PointData
        {
            public string X { get; set; }
            public double Y { get; set; }
        }

        /// <summary>
        /// 将RGB颜色值转换为16进制字符串（BGR转RGB）
        /// </summary>
        private static string ConvertRgbToHex(int rgbValue)
        {
            // Word API使用BGR格式，需要转换为RGB
            int b = (rgbValue >> 16) & 0xFF;
            int g = (rgbValue >> 8) & 0xFF;
            int r = rgbValue & 0xFF;
            return $"{r:X2}{g:X2}{b:X2}";
        }

        /// <summary>
        /// 转义XML属性值
        /// </summary>
        private static string EscapeXmlAttribute(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        /// <summary>
        /// 转义XML文本内容
        /// </summary>
        private static string EscapeXmlText(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            
            return value
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static void BeginReadImageContext(bool extractImages, string docPath, string imageSaveDir, string fileHash)
        {
            _readUseFullImages = extractImages
                && !string.IsNullOrEmpty(docPath)
                && !string.IsNullOrEmpty(imageSaveDir)
                && !string.IsNullOrEmpty(fileHash);
            _readDocPath = docPath;
            _readImageSaveDir = imageSaveDir;
            _readFileHash = fileHash;
            _readZip = null;
            _readRelsDict = null;
        }

        private static void EndReadImageContext()
        {
            _readZip?.Dispose();
            _readZip = null;
            _readRelsDict = null;
            _readUseFullImages = false;
            _readDocPath = null;
            _readImageSaveDir = null;
            _readFileHash = null;
        }

        private static string FormatImageTag(string imageId, string imagePath)
        {
            string idAttr = string.IsNullOrEmpty(imageId) ? "" : $" id=\"{EscapeXmlAttribute(imageId)}\"";
            string src = EscapeXmlAttribute(imagePath ?? "");
            return $"<image{idAttr} src=\"{src}\"/>\r";
        }

        private static string AllocatePlaceholderImagePath()
        {
            _placeholderImageCounter++;
            return $"lightweight/placeholder_{_placeholderImageCounter}.png";
        }

        private static void EnsureReadZipAndRels()
        {
            if (!_readUseFullImages || _readZip != null)
            {
                return;
            }

            _readZip = ZipFile.OpenRead(_readDocPath);
            _readRelsDict = new Dictionary<string, string>();
            var relsEntry = _readZip.GetEntry("word/_rels/document.xml.rels");
            if (relsEntry == null)
            {
                return;
            }

            try
            {
                using (var stream = relsEntry.Open())
                using (var reader = new StreamReader(stream))
                {
                    XDocument relsDoc = XDocument.Parse(reader.ReadToEnd());
                    XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                    foreach (var rel in relsDoc.Descendants(nsRels + "Relationship"))
                    {
                        string relId = rel.Attribute("Id")?.Value;
                        string target = rel.Attribute("Target")?.Value;
                        if (!string.IsNullOrEmpty(relId) && !string.IsNullOrEmpty(target))
                        {
                            _readRelsDict[relId] = target;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] 读取relationships文件失败: {ex.Message}");
            }
        }

        private static string ResolveImagePathForDrawing(XElement drawing)
        {
            if (drawing == null)
            {
                return AllocatePlaceholderImagePath();
            }

            XNamespace nsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            var blip = drawing.Descendants(nsA + "blip").FirstOrDefault();
            if (blip == null)
            {
                System.Diagnostics.Debug.WriteLine("[WordReader] ⚠️ drawing 未找到 blip，使用 placeholder");
                return AllocatePlaceholderImagePath();
            }

            if (!_readUseFullImages)
            {
                return AllocatePlaceholderImagePath();
            }

            EnsureReadZipAndRels();
            string embedAttr = blip.Attribute(nsR + "embed")?.Value;
            string linkAttr = blip.Attribute(nsR + "link")?.Value;
            string relId = embedAttr ?? linkAttr;
            if (string.IsNullOrEmpty(relId) || _readRelsDict == null || !_readRelsDict.TryGetValue(relId, out string target))
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 无法解析图片 relationship: {relId ?? "(null)"}，使用 placeholder");
                return AllocatePlaceholderImagePath();
            }

            string imagePathInZip;
            if (target.StartsWith("/"))
            {
                imagePathInZip = "word" + target;
            }
            else if (target.StartsWith("media/"))
            {
                imagePathInZip = "word/" + target;
            }
            else
            {
                imagePathInZip = "word/media/" + target;
            }

            var imageEntry = _readZip?.GetEntry(imagePathInZip);
            if (imageEntry == null)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 图片文件不存在于ZIP中: {imagePathInZip}，使用 placeholder");
                return AllocatePlaceholderImagePath();
            }

            try
            {
                string counterKey = _readFileHash.Length >= 10 ? _readFileHash.Substring(0, 10) : _readFileHash;
                if (!_imageCounters.ContainsKey(counterKey))
                {
                    _imageCounters[counterKey] = 0;
                }

                string ext = Path.GetExtension(imagePathInZip);
                if (string.IsNullOrEmpty(ext))
                {
                    ext = ".png";
                }

                _imageCounters[counterKey]++;
                string imageFilename = $"image_{_imageCounters[counterKey]}{ext}";
                string imageSavePath = Path.Combine(_readImageSaveDir, imageFilename);
                using (var sourceStream = imageEntry.Open())
                using (var targetStream = File.Create(imageSavePath))
                {
                    sourceStream.CopyTo(targetStream);
                }

                string hashDirName = Path.GetFileName(_readImageSaveDir);
                return $"{hashDirName}/{imageFilename}".Replace("\\", "/");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 单图落盘失败: {ex.Message}，使用 placeholder");
                return AllocatePlaceholderImagePath();
            }
        }

        /// <summary>
        /// 同 w:p 内：已有会读入的正文（w:ins / w:r 等）且其后仍有 w:del 时，段末不应追加 \r。
        /// del 被跳过后，¶ 若加在 ins 文字后会导致映射误带 ^p（与「文字+图」段内后继检测同类）。
        /// </summary>
        private static bool ParagraphHasReadableContentBeforeDeletionInSameParagraph(XElement paraElement, XNamespace w)
        {
            if (paraElement == null)
            {
                return false;
            }

            bool seenReadable = false;
            foreach (var child in paraElement.Elements())
            {
                if (IsRevisionDeletion(child, w))
                {
                    return seenReadable;
                }

                if (IsRevisionInsertion(child, w))
                {
                    seenReadable = true;
                    continue;
                }

                string localName = child.Name.LocalName;
                if (child.Name == w + "r" || child.Name == w + "hyperlink"
                    || child.Name == w + "drawing" || localName == "AlternateContent")
                {
                    seenReadable = true;
                }
            }

            return false;
        }

        /// <summary>
        /// 段落仅含 w:del（及 pPr 等），读入后 paraContent 为空时，不应再追加 \r。
        /// 否则幽灵段界会挂在上一段末句（如 S_0006i）上。
        /// </summary>
        private static bool ShouldSuppressEmptyDeletionOnlyParagraphR(XElement paraElement, XNamespace w)
        {
            if (paraElement == null)
            {
                return false;
            }

            bool hasDeletion = false;
            foreach (var child in paraElement.Elements())
            {
                string localName = child.Name.LocalName;
                if (localName.EndsWith("Pr", StringComparison.Ordinal))
                {
                    if (child.Descendants().Any(el => IsRevisionDeletion(el, w)))
                    {
                        hasDeletion = true;
                    }

                    continue;
                }

                if (localName == "proofErr" || localName.StartsWith("bookmark", StringComparison.Ordinal))
                {
                    continue;
                }

                if (IsRevisionDeletion(child, w))
                {
                    hasDeletion = true;
                    continue;
                }

                if (IsRevisionInsertion(child, w))
                {
                    return false;
                }

                if (child.Name == w + "r" || child.Name == w + "hyperlink"
                    || child.Name == w + "drawing" || localName == "AlternateContent")
                {
                    return false;
                }
            }

            return hasDeletion;
        }

        private static bool IsPageBreakBr(XElement el, XNamespace w)
        {
            if (el == null || el.Name != w + "br")
            {
                return false;
            }

            XName typeName = w + "type";
            string typeVal = (string)el.Attribute(typeName);
            if (string.IsNullOrEmpty(typeVal))
            {
                typeVal = el.Attributes()
                    .FirstOrDefault(a => a.Name.LocalName == "type")
                    ?.Value;
            }

            return string.Equals(typeVal, "page", StringComparison.Ordinal);
        }

        private static bool IsLayoutOnlyBreakElement(XElement el, XNamespace w)
        {
            if (el == null)
            {
                return false;
            }

            if (el.Name.LocalName == "lastRenderedPageBreak")
            {
                return true;
            }

            return IsPageBreakBr(el, w);
        }

        /// <summary>
        /// 分页符/分节符等纯排版段落读入为空时，不应再追加 \r。
        /// </summary>
        private static bool ShouldSuppressLayoutOnlyParagraphR(XElement paraElement, XNamespace w)
        {
            if (paraElement == null)
            {
                return false;
            }

            foreach (var tEl in paraElement.Descendants(w + "t"))
            {
                if (!string.IsNullOrWhiteSpace(tEl.Value))
                {
                    return false;
                }
            }

            // 仅有脚注/尾注引用也算有正文内容（ReadText 会输出 ^f/^e）
            if (paraElement.Descendants(w + "footnoteReference").Any()
                || paraElement.Descendants(w + "endnoteReference").Any())
            {
                return false;
            }

            foreach (var brEl in paraElement.Descendants(w + "br"))
            {
                if (!IsPageBreakBr(brEl, w))
                {
                    return false;
                }
            }

            var pPr = paraElement.Element(w + "pPr");
            if (pPr?.Element(w + "sectPr") != null)
            {
                return true;
            }

            foreach (var el in paraElement.Descendants())
            {
                if (IsLayoutOnlyBreakElement(el, w))
                {
                    continue;
                }

                string localName = el.Name.LocalName;
                if (localName == "p" || localName == "pPr" || localName == "r" || localName == "rPr"
                    || localName == "proofErr" || localName.StartsWith("bookmark", StringComparison.Ordinal))
                {
                    continue;
                }

                if (localName.EndsWith("Pr", StringComparison.Ordinal))
                {
                    continue;
                }

                if (localName == "t")
                {
                    continue;
                }

                if (IsRevisionDeletion(el, w))
                {
                    continue;
                }

                return false;
            }

            return paraElement.Descendants().Any(el => IsLayoutOnlyBreakElement(el, w));
        }

        private static void AppendParagraphToResult(
            ref string result,
            string paraContent,
            bool hasChart,
            bool suppressTrailingR = false,
            bool skipEmptyParagraphR = false)
        {
            if (string.IsNullOrEmpty(paraContent))
            {
                if (!hasChart && !skipEmptyParagraphR)
                {
                    result += "\r";
                }
                return;
            }

            if (paraContent.EndsWith("\r", StringComparison.Ordinal) || suppressTrailingR)
            {
                result += paraContent;
            }
            else
            {
                result += paraContent + "\r";
            }
        }

        private static void ProcessParagraphNodesInOrder(
            IEnumerable<XNode> nodes,
            XNamespace w,
            StringBuilder textBuffer,
            StringBuilder output,
            HashSet<XElement> emittedDrawings,
            List<string> orderLog)
        {
            foreach (var node in nodes)
            {
                if (!(node is XElement el))
                {
                    continue;
                }

                if (el.Name == w + "t")
                {
                    string val = el.Value ?? "";
                    if (val.Length > 0)
                    {
                        textBuffer.Append(val);
                        orderLog.Add($"t({val.Length})");
                    }
                }
                else if (el.Name == w + "footnoteReference")
                {
                    // 与 Word Find 特殊码一致，便于后续按映射串定位
                    textBuffer.Append("^f");
                    orderLog.Add("footnoteReference");
                }
                else if (el.Name == w + "endnoteReference")
                {
                    textBuffer.Append("^e");
                    orderLog.Add("endnoteReference");
                }
                else if (el.Name == w + "tab")
                {
                    textBuffer.Append("\t");
                    orderLog.Add("tab");
                }
                else if (el.Name == w + "br")
                {
                    if (IsPageBreakBr(el, w))
                    {
                        orderLog.Add("br:page_skip");
                    }
                    else
                    {
                        textBuffer.Append("\n");
                        orderLog.Add("br");
                    }
                }
                else if (IsLayoutOnlyBreakElement(el, w))
                {
                    orderLog.Add("layout_break_skip");
                }
                else if (el.Name == w + "drawing")
                {
                    if (emittedDrawings.Add(el))
                    {
                        if (textBuffer.Length > 0)
                        {
                            output.Append(System.Security.SecurityElement.Escape(textBuffer.ToString()));
                            textBuffer.Clear();
                        }

                        output.Append(AppendDrawingContent(el, w));
                        orderLog.Add(IsChartDrawing(el) ? "chart-drawing" : "drawing");
                    }
                }
                else if (IsRevisionInsertion(el, w))
                {
                    orderLog.Add("ins");
                    ProcessParagraphNodesInOrder(el.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                }
                else if (IsRevisionDeletion(el, w))
                {
                    orderLog.Add("del_skip");
                }
                else
                {
                    ProcessParagraphNodesInOrder(el.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                }
            }
        }

        private static string BuildParagraphContentInOrder(XElement paraElement, XNamespace w)
        {
            if (paraElement == null)
            {
                return "";
            }

            var textBuffer = new StringBuilder();
            var output = new StringBuilder();
            var emittedDrawings = new HashSet<XElement>();
            var orderLog = new List<string>();

            foreach (var child in paraElement.Elements())
            {
                string localName = child.Name.LocalName;
                if (child.Name == w + "r" || child.Name == w + "hyperlink")
                {
                    orderLog.Add(child.Name == w + "r" ? "r" : "hyperlink");
                    ProcessParagraphNodesInOrder(child.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                }
                else if (child.Name == w + "br")
                {
                    if (IsPageBreakBr(child, w))
                    {
                        orderLog.Add("p:br:page_skip");
                    }
                    else
                    {
                        textBuffer.Append("\n");
                        orderLog.Add("p:br");
                    }
                }
                else if (IsLayoutOnlyBreakElement(child, w))
                {
                    orderLog.Add("p:layout_break_skip");
                }
                else if (child.Name == w + "drawing")
                {
                    if (emittedDrawings.Add(child))
                    {
                        if (textBuffer.Length > 0)
                        {
                            output.Append(System.Security.SecurityElement.Escape(textBuffer.ToString()));
                            textBuffer.Clear();
                        }

                        output.Append(AppendDrawingContent(child, w));
                        orderLog.Add(IsChartDrawing(child) ? "p:chart-drawing" : "p:drawing");
                    }
                }
                else if (localName == "AlternateContent")
                {
                    XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
                    var choice = child.Element(mc + "Choice");
                    if (choice != null)
                    {
                        ProcessParagraphNodesInOrder(choice.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                    }
                }
                else if (IsRevisionInsertion(child, w))
                {
                    orderLog.Add("p:ins");
                    ProcessParagraphNodesInOrder(child.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                }
                else if (IsRevisionDeletion(child, w))
                {
                    orderLog.Add("p:del_skip");
                }
                else if (!localName.EndsWith("Pr"))
                {
                    ProcessParagraphNodesInOrder(child.Nodes(), w, textBuffer, output, emittedDrawings, orderLog);
                }
            }

            if (textBuffer.Length > 0)
            {
                output.Append(System.Security.SecurityElement.Escape(textBuffer.ToString()));
            }

            string paraResult = output.ToString();
            if (paraResult.Contains(DebugParagraphKeyword))
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] 段内顺序: {string.Join(", ", orderLog)}");
                System.Diagnostics.Debug.WriteLine(
                    $"[WordReader] 段落最终内容（包含图片）: {paraResult.Substring(0, Math.Min(200, paraResult.Length))}...");
            }

            return paraResult;
        }

        /// <summary>
        /// 递归处理 XML 元素，一步到位提取所有内容
        /// </summary>
        /// <param name="element">要处理的 XML 元素</param>
        /// <param name="w">Word OpenXML 命名空间</param>
        /// <param name="document">Word 文档对象</param>
        /// <param name="docPath">文档路径</param>
        /// <param name="imageSaveDir">图片保存目录</param>
        /// <param name="fileHash">文件hash值</param>
        /// <param name="chartDataMap">图表数据映射（位置 -> 图表XML）</param>
        /// <returns>处理后的内容字符串</returns>
        private static string ProcessElementRecursive(XElement element, XNamespace w, Word.Document document, string docPath = null, string imageSaveDir = null, string fileHash = null)
        {
            string result = "";

            try
            {
                // 遍历当前元素的所有直接子元素
                foreach (var child in element.Elements())
                {
                    // 处理段落 w:p
                    if (child.Name == w + "p")
                    {
                        string paraContent = BuildParagraphContentInOrder(child, w);
                        bool hasChart = paraContent != null
                            && paraContent.IndexOf("<Chart", StringComparison.OrdinalIgnoreCase) >= 0;
                        if (!string.IsNullOrEmpty(paraContent)
                            && paraContent.StartsWith("<image", StringComparison.Ordinal)
                            && paraContent.IndexOf("<image", 1, StringComparison.Ordinal) < 0)
                        {
                            int textAfterImage = paraContent.IndexOf('\r');
                            bool imageOnly = textAfterImage >= 0
                                && textAfterImage + 1 >= paraContent.Length;
                            if (imageOnly)
                            {
                                EasyWriteDiagnostics.LogDocumentExtractVerbose(
                                    $"[WordReader] 段落只有图片，内容: {paraContent.Substring(0, Math.Min(paraContent.Length, textAfterImage + 1))}");
                            }
                        }

                        bool suppressTrailingR = ParagraphHasReadableContentBeforeDeletionInSameParagraph(child, w);
                        if (suppressTrailingR && !string.IsNullOrEmpty(paraContent))
                        {
                            EasyWriteDiagnostics.LogDocumentExtractVerbose(
                                "[WordReader][revision] 同段 ins 后有 w:del，段末不追加 \\r");
                        }

                        bool skipEmptyParagraphR = string.IsNullOrEmpty(paraContent)
                            && ShouldSuppressEmptyDeletionOnlyParagraphR(child, w);
                        if (skipEmptyParagraphR)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                "[WordReader][revision] 仅含 w:del 的段落读入为空，不追加 \\r");
                        }

                        bool skipLayoutOnlyParagraphR = string.IsNullOrEmpty(paraContent)
                            && ShouldSuppressLayoutOnlyParagraphR(child, w);

                        AppendParagraphToResult(
                            ref result, paraContent, hasChart, suppressTrailingR,
                            skipEmptyParagraphR || skipLayoutOnlyParagraphR);
                    }
                    // 处理表格 w:tbl
                    else if (child.Name == w + "tbl")
                    {
                        // 创建表格级别的合并信息跟踪
                        var mergeContext = new TableMergeContext();

                        string tableContent = ProcessTableRecursive(child, w, document, mergeContext, docPath, imageSaveDir, fileHash);
                        string tableId = TakeWordSourcedTableId(tableContent, child, w);
                        result += $"<table id=\"{tableId}\">";
                        result += tableContent;
                        result += "</table>\n";
                    }
                    // 块级修订：w:ins 内可能含 w:p，纳入读回；w:del 为已删内容，跳过
                    else if (IsRevisionInsertion(child, w))
                    {
                        result += ProcessElementRecursive(child, w, document, docPath, imageSaveDir, fileHash);
                    }
                    else if (IsRevisionDeletion(child, w))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[WordReader][revision] skip block {child.Name.LocalName}");
                    }
                    // 其他元素（如属性元素），跳过
                    else
                    {
                        // 跳过属性元素（pPr, rPr, tcPr, tblPr 等）
                        if (!child.Name.LocalName.EndsWith("Pr"))
                        {
                            // 如果不是属性元素，继续递归处理
                            result += ProcessElementRecursive(child, w, document, docPath, imageSaveDir, fileHash);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogExtractVerbose($"[DEBUG] 递归处理元素失败: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// 专门处理表格，提取合并信息
        /// </summary>
        /// <param name="tableElement">表格元素</param>
        /// <param name="w">Word OpenXML 命名空间</param>
        /// <param name="document">Word 文档对象</param>
        /// <param name="mergeContext">合并信息跟踪上下文</param>
        /// <param name="docPath">文档路径</param>
        /// <param name="imageSaveDir">图片保存目录</param>
        /// <param name="fileHash">文件hash值</param>
        /// <param name="chartDataMap">图表数据映射（位置 -> 图表XML）</param>
        /// <returns>处理后的内容字符串</returns>
        private static string ProcessTableRecursive(XElement tableElement, XNamespace w, Word.Document document, TableMergeContext mergeContext, string docPath = null, string imageSaveDir = null, string fileHash = null)
        {
            string result = "";
            
            try
            {
                // 获取所有行
                var rows = tableElement.Elements(w + "tr").ToList();
                LogExtractVerbose($"[DEBUG] 开始处理表格，共 {rows.Count} 行");
                
                // 预先扫描表格，计算所有纵向合并的 rowspan
                var rowspanMap = CalculateAllRowspans(rows, w, mergeContext);
                LogExtractVerbose($"[DEBUG] 扫描完成，找到 {rowspanMap.Count} 个纵向合并单元格");
                
                // 处理每一行
                foreach (var row in rows)
                {
                    mergeContext.CurrentRowIndex++;  // 行索引递增
                    mergeContext.CurrentLogicalColumnIndex = 0;  // 重置列索引
                    
                    var cells = row.Elements(w + "tc").ToList();
                    LogExtractVerbose($"[DEBUG] 处理第 {mergeContext.CurrentRowIndex + 1} 行，共 {cells.Count} 个单元格");
                    
                    result += "<row>";  // 添加行开始标签
                    
                    // 处理该行的所有单元格
                    foreach (var cell in cells)
                    {
                        result += ProcessCellWithMerge(cell, w, document, mergeContext, rowspanMap, docPath, imageSaveDir, fileHash);
                    }
                    
                    result += "</row>\n";  // 闭合行标签，添加换行符
                }
            }
            catch (Exception ex)
            {
                LogExtractVerbose($"[DEBUG] 表格处理失败: {ex.Message}");
                LogExtractVerbose($"[DEBUG] 异常堆栈: {ex.StackTrace}");
            }
            
            return result;
        }

        /// <summary>
        /// 处理单元格，提取合并信息并递归提取内容
        /// </summary>
        /// <param name="cell">单元格元素</param>
        /// <param name="w">Word OpenXML 命名空间</param>
        /// <param name="document">Word 文档对象</param>
        /// <param name="mergeContext">合并信息跟踪上下文</param>
        /// <param name="rowspanMap">rowspan映射表：key = (行索引, 列索引), value = rowspan值</param>
        /// <param name="docPath">文档路径</param>
        /// <param name="imageSaveDir">图片保存目录</param>
        /// <param name="fileHash">文件hash值</param>
        /// <param name="chartDataMap">图表数据映射（位置 -> 图表XML）</param>
        /// <returns>处理后的单元格字符串</returns>
        private static string ProcessCellWithMerge(XElement cell, XNamespace w, Word.Document document, TableMergeContext mergeContext, Dictionary<(int row, int col), int> rowspanMap, string docPath = null, string imageSaveDir = null, string fileHash = null)
        {
            string result = "";
            
            try
            {
                var tcPr = cell.Element(w + "tcPr");
                
                // 调试：检查 tcPr 是否存在
                if (tcPr == null)
                {
                    LogExtractVerbose($"[DEBUG] 警告: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex} 的单元格没有 tcPr 元素");
                }
                else
                {
                    // 打印 tcPr 的所有子元素名称
                    var tcPrChildren = tcPr.Elements().Select(e => e.Name.LocalName).ToList();
                    LogExtractVerbose($"[DEBUG] 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex} 的 tcPr 子元素: {string.Join(", ", tcPrChildren)}");
                }
                
                // 1. 提取 gridSpan（横向合并）
                int colspan = 1;
                var gridSpan = tcPr?.Element(w + "gridSpan");
                if (gridSpan != null)
                {
                    string gridSpanVal = GetValAttribute(gridSpan);
                    LogExtractVerbose($"[DEBUG] 检测到 gridSpan 元素: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, val={gridSpanVal ?? "(null)"}");
                    
                    if (!string.IsNullOrEmpty(gridSpanVal) && int.TryParse(gridSpanVal, out int parsedColspan))
                    {
                        colspan = parsedColspan;
                        if (colspan > 1)
                        {
                            LogExtractVerbose($"[DEBUG] ✅ 检测到横向合并: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, colspan={colspan}");
                        }
                    }
                }
                
                // 2. 提取 vMerge（纵向合并）
                bool isVerticalMergeContinue = false;
                int rowspan = 1;
                
                var vMerge = tcPr?.Element(w + "vMerge");
                if (vMerge != null)
                {
                    string vMergeVal = GetValAttribute(vMerge);
                    LogExtractVerbose($"[DEBUG] 检测到vMerge: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, val={vMergeVal ?? "(null/延续)"}");
                    
                    if (vMergeVal == "restart")
                    {
                        // 纵向合并起始
                        // 记录合并起始信息
                        mergeContext.VerticalMerges[mergeContext.CurrentLogicalColumnIndex] = 
                            (mergeContext.CurrentRowIndex, colspan);
                        
                        // 从rowspanMap获取rowspan值
                        if (rowspanMap.TryGetValue((mergeContext.CurrentRowIndex, mergeContext.CurrentLogicalColumnIndex), out int calculatedRowspan))
                        {
                            rowspan = calculatedRowspan;
                            LogExtractVerbose($"[DEBUG] ✅ 纵向合并起始: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, rowspan={rowspan}");
                        }
                        else
                        {
                            LogExtractVerbose($"[DEBUG] ⚠️ 纵向合并起始但未在rowspanMap中找到: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}");
                        }
                    }
                    else
                    {
                        // 纵向合并延续（val 为空或不存在）
                        isVerticalMergeContinue = true;
                        // 从rowspanMap获取rowspan值（延续单元格的rowspan应该与起始单元格相同）
                        if (rowspanMap.TryGetValue((mergeContext.CurrentRowIndex, mergeContext.CurrentLogicalColumnIndex), out int calculatedRowspan))
                        {
                            rowspan = calculatedRowspan;
                            LogExtractVerbose($"[DEBUG] ✅ 纵向合并延续: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, rowspan={rowspan}");
                        }
                        else
                        {
                            LogExtractVerbose($"[DEBUG] ⚠️ 纵向合并延续但未在rowspanMap中找到: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}");
                        }
                    }
                }
                
                // 3. 如果是纵向合并的延续单元格，按照HTML格式，不输出（rowspan已经表示占据多行）
                if (isVerticalMergeContinue)
                {
                    LogExtractVerbose($"[DEBUG] 跳过延续单元格输出（HTML格式）: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}");
                    // 更新逻辑列索引（重要！即使不输出也要更新）
                    mergeContext.CurrentLogicalColumnIndex += colspan;
                    return "";  // 返回空字符串，不输出延续单元格
                }
                
                // 4. 构建 cell 标签属性（只对起始单元格和普通单元格）
                var cellAttributes = new List<string>();
                if (colspan > 1)
                {
                    cellAttributes.Add($"colspan=\"{colspan}\"");
                }
                if (rowspan > 1)
                {
                    cellAttributes.Add($"rowspan=\"{rowspan}\"");
                }
                
                string attrStr = cellAttributes.Count > 0 
                    ? " " + string.Join(" ", cellAttributes) 
                    : "";
                
                // 调试输出：如果有合并信息，打印出来
                if (cellAttributes.Count > 0)
                {
                    LogExtractVerbose($"[DEBUG] 单元格合并属性: 行{mergeContext.CurrentRowIndex}, 列{mergeContext.CurrentLogicalColumnIndex}, 属性={attrStr}");
                }
                
                // 5. 递归提取单元格内容（段落、文本、嵌套表格等）
                string cellContent = ProcessElementRecursive(cell, w, document, docPath, imageSaveDir, fileHash);
                
                // 5.1 去掉单元格内容最后的换行符（\r）
                // 单元格内最后一个换行符只是表示单元格内是一个段落，不是实际的换行符
                cellContent = cellContent.TrimEnd('\r');
                
                // 5.2 如果单元格内容以 <image /> 结尾，需要添加 \r（统一格式）
                // 匹配 <image /> 或 <image ... /> 格式（支持带属性的自闭合标签）
                if (Regex.IsMatch(cellContent, @"<image\s+[^>]*/>\s*$", RegexOptions.IgnoreCase))
                {
                    cellContent += "\r";
                }
                
                // 6. 输出带合并信息的 cell 标签
                result = $"<cell{attrStr}>";
                result += cellContent;  // 递归提取的内容
                result += "</cell>\n";
                
                // 6. 更新逻辑列索引（重要！必须加上colspan）
                mergeContext.CurrentLogicalColumnIndex += colspan;
            }
            catch (Exception ex)
            {
                LogExtractVerbose($"[DEBUG] 单元格处理失败: {ex.Message}");
                // 降级处理：输出基本cell标签
                string cellContent = ProcessElementRecursive(cell, w, document, docPath, imageSaveDir, fileHash);
                // 去掉单元格内容最后的换行符（\r）
                cellContent = cellContent.TrimEnd('\r');
                // 如果单元格内容以 <image /> 结尾，需要添加 \r（统一格式）
                if (Regex.IsMatch(cellContent, @"<image\s+[^>]*/>\s*$", RegexOptions.IgnoreCase))
                {
                    cellContent += "\r";
                }
                result = $"<cell>{cellContent}</cell>\n";
            }
            
            return result;
        }

        /// <summary>
        /// 获取元素的 val 属性值（支持命名空间）
        /// </summary>
        private static string GetValAttribute(XElement element)
        {
            if (element == null) return null;
            
            // 方法1：直接获取 val 属性
            var valAttr = element.Attribute("val");
            if (valAttr != null)
            {
                return valAttr.Value;
            }
            
            // 方法2：尝试带命名空间的属性
            var ns = element.Name.Namespace;
            var valAttrWithNs = element.Attribute(ns + "val");
            if (valAttrWithNs != null)
            {
                return valAttrWithNs.Value;
            }
            
            return null;
        }

        /// <summary>
        /// 预先扫描表格，计算所有纵向合并的rowspan值
        /// </summary>
        /// <param name="rows">所有行元素</param>
        /// <param name="w">Word OpenXML 命名空间</param>
        /// <param name="mergeContext">合并信息跟踪上下文（用于重置）</param>
        /// <returns>rowspan映射表：key = (行索引, 列索引), value = rowspan值</returns>
        private static Dictionary<(int row, int col), int> CalculateAllRowspans(List<XElement> rows, XNamespace w, TableMergeContext mergeContext)
        {
            var rowspanMap = new Dictionary<(int row, int col), int>();
            // 改为列表，存储所有纵向合并起始位置（允许同一列有多个起始位置）
            var verticalMergeStarts = new List<(int startRow, int startCol, int gridSpan)>();
            
            // 第一遍扫描：找到所有纵向合并的起始位置
            for (int rowIdx = 0; rowIdx < rows.Count; rowIdx++)
            {
                var row = rows[rowIdx];
                int logicalColIdx = 0;
                
                foreach (var cell in row.Elements(w + "tc"))
                {
                    var tcPr = cell.Element(w + "tcPr");
                    
                    // 获取gridSpan
                    int gridSpan = 1;
                    var gridSpanElem = tcPr?.Element(w + "gridSpan");
                    if (gridSpanElem != null)
                    {
                        string gridSpanVal = GetValAttribute(gridSpanElem);
                        if (!string.IsNullOrEmpty(gridSpanVal))
                        {
                            int.TryParse(gridSpanVal, out gridSpan);
                        }
                    }
                    
                    // 检查vMerge
                    var vMerge = tcPr?.Element(w + "vMerge");
                    if (vMerge != null)
                    {
                        string vMergeVal = GetValAttribute(vMerge);
                        if (vMergeVal == "restart")
                        {
                            // 记录纵向合并起始（使用列表存储，避免同一列的多个起始被覆盖）
                            verticalMergeStarts.Add((rowIdx, logicalColIdx, gridSpan));
                            LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 找到纵向合并起始 - 行{rowIdx}, 列{logicalColIdx}, gridSpan={gridSpan}");
                        }
                    }
                    
                    logicalColIdx += gridSpan;
                }
            }
            
            LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 找到 {verticalMergeStarts.Count} 个纵向合并起始位置");
            
            // 第二遍扫描：计算每个纵向合并的rowspan
            foreach (var (startRow, startCol, gridSpan) in verticalMergeStarts)
            {
                // 从起始行开始，向下查找延续单元格
                int rowspan = 1;
                LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 开始计算 rowspan，起始行={startRow}, 列={startCol}, gridSpan={gridSpan}");
                for (int rowIdx = startRow + 1; rowIdx < rows.Count; rowIdx++)
                {
                    var row = rows[rowIdx];
                    int logicalColIdx = 0;
                    bool foundContinue = false;
                    
                    foreach (var cell in row.Elements(w + "tc"))
                    {
                        var tcPr = cell.Element(w + "tcPr");
                        
                        // 获取gridSpan
                        int cellGridSpan = 1;
                        var gridSpanElem = tcPr?.Element(w + "gridSpan");
                        if (gridSpanElem != null)
                        {
                            string gridSpanVal = GetValAttribute(gridSpanElem);
                            if (!string.IsNullOrEmpty(gridSpanVal))
                            {
                                int.TryParse(gridSpanVal, out cellGridSpan);
                            }
                        }
                        
                        // 检查是否是对应列的延续单元格
                        if (logicalColIdx == startCol)
                        {
                            var vMerge = tcPr?.Element(w + "vMerge");
                            if (vMerge != null)
                            {
                                string vMergeVal = GetValAttribute(vMerge);
                                LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 检查延续单元格，行={rowIdx}, 列={logicalColIdx}, vMerge val={vMergeVal ?? "(null)"}");
                                if (vMergeVal != "restart")  // 延续单元格：val 为空或不存在
                                {
                                    // 找到延续单元格
                                    foundContinue = true;
                                    rowspan++;
                                    LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 找到延续单元格，行={rowIdx}, rowspan={rowspan}");
                                    
                                    // 验证gridSpan是否一致
                                    if (cellGridSpan != gridSpan)
                                    {
                                        LogExtractVerbose($"[DEBUG] 警告：纵向合并延续单元格的gridSpan不一致。起始: {gridSpan}, 延续: {cellGridSpan}");
                                    }
                                    break;
                                }
                                else if (vMergeVal == "restart")
                                {
                                    // 遇到新的起始单元格，合并结束
                                    LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 遇到新的起始单元格，合并结束，行={rowIdx}");
                                    foundContinue = false;
                                    break;
                                }
                            }
                        }
                        
                        logicalColIdx += cellGridSpan;
                    }
                    
                    if (!foundContinue)
                    {
                        // 没有找到延续单元格，合并结束
                        LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 未找到延续单元格，合并结束，最终 rowspan={rowspan}");
                        break;
                    }
                }
                
                // 记录rowspan值（起始单元格和所有延续单元格都使用相同的rowspan）
                LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 记录 rowspan，起始行={startRow}, 列={startCol}, rowspan={rowspan}");
                for (int r = startRow; r < startRow + rowspan; r++)
                {
                    rowspanMap[(r, startCol)] = rowspan;
                    LogExtractVerbose($"[DEBUG] CalculateAllRowspans: 记录 rowspan_map[({r}, {startCol})] = {rowspan}");
                }
            }
            
            return rowspanMap;
        }

        /// <summary>
        /// 计算相对路径（兼容 .NET Framework）
        /// </summary>
        /// <param name="fromPath">基础路径</param>
        /// <param name="toPath">目标路径</param>
        /// <returns>相对路径</returns>
        private static string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath) || string.IsNullOrEmpty(toPath))
            {
                return toPath;
            }

            // 规范化路径
            fromPath = Path.GetFullPath(fromPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            toPath = Path.GetFullPath(toPath).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            // 如果路径相同，返回当前目录
            if (fromPath.Equals(toPath, StringComparison.OrdinalIgnoreCase))
            {
                return ".";
            }

            // 检查是否在同一驱动器
            if (Path.GetPathRoot(fromPath) != Path.GetPathRoot(toPath))
            {
                return toPath; // 不同驱动器，返回绝对路径
            }

            // 找到共同的基础路径
            var fromParts = fromPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var toParts = toPath.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            int commonLength = 0;
            int minLength = Math.Min(fromParts.Length, toParts.Length);
            for (int i = 0; i < minLength; i++)
            {
                if (fromParts[i].Equals(toParts[i], StringComparison.OrdinalIgnoreCase))
                {
                    commonLength++;
                }
                else
                {
                    break;
                }
            }

            // 构建相对路径
            var relativeParts = new List<string>();
            
            // 添加 .. 回到共同基础
            for (int i = commonLength; i < fromParts.Length; i++)
            {
                relativeParts.Add("..");
            }
            
            // 添加目标路径的剩余部分
            for (int i = commonLength; i < toParts.Length; i++)
            {
                relativeParts.Add(toParts[i]);
            }

            return relativeParts.Count > 0 ? string.Join(Path.DirectorySeparatorChar.ToString(), relativeParts) : ".";
        }

        /// <summary>
        /// 计算文件的SHA-256哈希值（与Python代码对齐）
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>文件的SHA-256哈希值（十六进制字符串）</returns>
        private static string CalculateFileHash(string filePath)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            try
            {
                using (var sha256 = SHA256.Create())
                using (var fileStream = File.OpenRead(filePath))
                {
                    // 分块读取文件，避免大文件占用过多内存（与Python代码对齐，使用4096字节块）
                    byte[] buffer = new byte[4096];
                    int bytesRead;
                    while ((bytesRead = fileStream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        sha256.TransformBlock(buffer, 0, bytesRead, null, 0);
                    }
                    // 完成哈希计算（传递空数组表示没有更多数据）
                    sha256.TransformFinalBlock(new byte[0], 0, 0);
                    
                    // 转换为十六进制字符串（与Python的hexdigest()对齐）
                    StringBuilder sb = new StringBuilder();
                    foreach (byte b in sha256.Hash)
                    {
                        sb.Append(b.ToString("x2"));
                    }
                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] 计算文件hash失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取图片保存目录
        /// 目录结构：{文件所在目录}\{hash前10位}
        /// 与Python代码对齐：使用文件hash前10位作为文件夹名
        /// </summary>
        /// <param name="docPath">文档路径</param>
        /// <param name="fileHash">文件hash值</param>
        /// <returns>图片保存目录的完整路径</returns>
        private static string GetImageSaveDirectory(string docPath, string fileHash)
        {
            if (string.IsNullOrEmpty(docPath) || string.IsNullOrEmpty(fileHash))
            {
                return null;
            }

            try
            {
                // 获取文件所在目录
                string fileDir = Path.GetDirectoryName(docPath);
                if (string.IsNullOrEmpty(fileDir))
                {
                    System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 无法获取文件目录");
                    return null;
                }
                
                // 使用hash前10位作为文件夹名（与Python代码对齐）
                string hashDirName = fileHash.Length >= 10 ? fileHash.Substring(0, 10) : fileHash;
                string imageDir = Path.Combine(fileDir, hashDirName);
                
                // 创建目录（如果不存在）
                Directory.CreateDirectory(imageDir);
                
                System.Diagnostics.Debug.WriteLine($"[WordReader] 文件目录: {fileDir}");
                System.Diagnostics.Debug.WriteLine($"[WordReader] hash文件夹名: {hashDirName}");
                System.Diagnostics.Debug.WriteLine($"[WordReader] 图片保存目录（完整路径）: {imageDir}");
                
                // 返回完整路径（用于 File.Create 等操作）
                return imageDir;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] 创建图片保存目录失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 轻量读入：不打开 docx zip，为每个 blip/drawing 生成占位路径，保证 &lt;image&gt; 标签与段落边界与完整模式一致。
        /// </summary>
        private static List<string> ExtractImagePlaceholderPaths(XElement paraElement)
        {
            var paths = new List<string>();
            if (paraElement == null)
            {
                return paths;
            }

            var blips = CollectBlipsFromParagraph(paraElement, out int drawingCount);
            int imageCount = blips.Count > 0 ? blips.Count : drawingCount;
            for (int i = 0; i < imageCount; i++)
            {
                _placeholderImageCounter++;
                paths.Add($"lightweight/placeholder_{_placeholderImageCounter}.png");
            }

            return paths;
        }

        private static List<XElement> CollectBlipsFromParagraph(XElement paraElement, out int drawingCount)
        {
            XNamespace nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
            XNamespace nsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
            XNamespace nsPic = "http://schemas.openxmlformats.org/drawingml/2006/picture";

            var blips = new List<XElement>();
            blips.AddRange(paraElement.Descendants(nsA + "blip"));

            foreach (var drawing in paraElement.Descendants(nsW + "drawing"))
            {
                foreach (var blip in drawing.Descendants(nsA + "blip"))
                {
                    if (!blips.Contains(blip))
                    {
                        blips.Add(blip);
                    }
                }
            }

            foreach (var pic in paraElement.Descendants(nsPic + "pic"))
            {
                foreach (var blip in pic.Descendants(nsA + "blip"))
                {
                    if (!blips.Contains(blip))
                    {
                        blips.Add(blip);
                    }
                }
            }

            foreach (var run in paraElement.Descendants(nsW + "r"))
            {
                foreach (var drawing in run.Descendants(nsW + "drawing"))
                {
                    foreach (var blip in drawing.Descendants(nsA + "blip"))
                    {
                        if (!blips.Contains(blip))
                        {
                            blips.Add(blip);
                        }
                    }
                }
            }

            drawingCount = paraElement.Descendants(nsW + "drawing").Count();
            return blips;
        }

        /// <summary>
        /// 从段落元素中提取图片并保存到指定目录
        /// </summary>
        private static List<string> ExtractAndSaveImages(string docPath, XElement paraElement, string imageSaveDir, string fileHash)
        {
            List<string> imagePaths = new List<string>();
            
            if (string.IsNullOrEmpty(docPath) || paraElement == null || string.IsNullOrEmpty(imageSaveDir) || string.IsNullOrEmpty(fileHash))
            {
                return imagePaths;
            }

            try
            {
                XNamespace nsW = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                XNamespace nsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
                var blips = CollectBlipsFromParagraph(paraElement, out int drawingCount);
                var allDrawings = paraElement.Descendants(nsW + "drawing").ToList();
                System.Diagnostics.Debug.WriteLine($"[WordReader] 段落中找到 {allDrawings.Count} 个drawing元素, {blips.Count} 个blip元素");

                if (blips.Count == 0)
                {
                    if (allDrawings.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 找到 {allDrawings.Count} 个drawing元素但未找到blip，可能图片格式不同");
                        foreach (var drawing in allDrawings.Take(2))
                        {
                            try
                            {
                                string drawingXml = drawing.ToString();
                                System.Diagnostics.Debug.WriteLine($"[WordReader] Drawing XML: {drawingXml.Substring(0, Math.Min(200, drawingXml.Length))}...");
                            }
                            catch { }
                        }
                    }
                    return imagePaths;
                }

                System.Diagnostics.Debug.WriteLine($"[WordReader] 段落中找到 {blips.Count} 个blip元素，开始提取图片");
                
                // 打开docx文件（本质是ZIP文件）
                using (ZipArchive zip = ZipFile.OpenRead(docPath))
                {
                    // 使用文件hash前10位作为计数器键（与Python代码对齐）
                    string counterKey = fileHash.Length >= 10 ? fileHash.Substring(0, 10) : fileHash;
                    if (!_imageCounters.ContainsKey(counterKey))
                    {
                        _imageCounters[counterKey] = 0;
                    }
                    
                    // 读取relationships文件
                    var relsEntry = zip.GetEntry("word/_rels/document.xml.rels");
                    Dictionary<string, string> relsDict = new Dictionary<string, string>();
                    
                    if (relsEntry != null)
                    {
                        try
                        {
                            using (var stream = relsEntry.Open())
                            using (var reader = new StreamReader(stream))
                            {
                                string relsXml = reader.ReadToEnd();
                                XDocument relsDoc = XDocument.Parse(relsXml);
                                XNamespace nsRels = "http://schemas.openxmlformats.org/package/2006/relationships";
                                
                                var relationships = relsDoc.Descendants(nsRels + "Relationship").ToList();
                                foreach (var rel in relationships)
                                {
                                    string relId = rel.Attribute("Id")?.Value;
                                    string target = rel.Attribute("Target")?.Value;
                                    if (!string.IsNullOrEmpty(relId) && !string.IsNullOrEmpty(target))
                                    {
                                        relsDict[relId] = target;
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"[WordReader] 读取relationships文件失败: {ex.Message}");
                        }
                    }
                    
                    foreach (var blip in blips)
                    {
                        // 获取图片的relationship ID
                        string embedAttr = blip.Attribute(nsR + "embed")?.Value;
                        string linkAttr = blip.Attribute(nsR + "link")?.Value;
                        
                        string relId = embedAttr ?? linkAttr;
                        if (string.IsNullOrEmpty(relId))
                        {
                            continue;
                        }
                        
                        // 从relationships字典中查找图片文件路径
                        if (relsDict.ContainsKey(relId))
                        {
                            string target = relsDict[relId];
                            
                            // 构建完整的图片路径
                            string imagePathInZip;
                            if (target.StartsWith("/"))
                            {
                                imagePathInZip = "word" + target;
                            }
                            else if (target.StartsWith("media/"))
                            {
                                imagePathInZip = "word/" + target;
                            }
                            else
                            {
                                imagePathInZip = "word/media/" + target;
                            }
                            
                            // 提取图片文件
                            var imageEntry = zip.GetEntry(imagePathInZip);
                            if (imageEntry != null)
                            {
                                // 获取图片文件扩展名
                                string ext = Path.GetExtension(imagePathInZip);
                                if (string.IsNullOrEmpty(ext))
                                {
                                    ext = ".png";  // 默认扩展名
                                }
                                
                                // 生成保存的文件名（使用递增计数器）
                                _imageCounters[counterKey]++;
                                int imageCounter = _imageCounters[counterKey];
                                string imageFilename = $"image_{imageCounter}{ext}";
                                string imageSavePath = Path.Combine(imageSaveDir, imageFilename);
                                
                                // 保存图片
                                using (var sourceStream = imageEntry.Open())
                                using (var targetStream = File.Create(imageSavePath))
                                {
                                    sourceStream.CopyTo(targetStream);
                                }
                                
                                // 记录相对路径：{hash前10位}/{imageFilename}（与Python代码对齐）
                                // 从 imageSaveDir 中提取 hash 文件夹名（最后一个目录名）
                                string hashDirName = Path.GetFileName(imageSaveDir);
                                string relativePath = $"{hashDirName}/{imageFilename}".Replace("\\", "/");
                                imagePaths.Add(relativePath);
                                System.Diagnostics.Debug.WriteLine($"[WordReader] ✅ 提取图片成功: 相对路径={relativePath}, 保存路径={imageSavePath}");
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 图片文件不存在于ZIP中: {imagePathInZip}");
                            }
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"[WordReader] ⚠️ 未找到relationship ID: {relId}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[WordReader] 图片提取过程出错: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[WordReader] 异常堆栈: {ex.StackTrace}");
            }
            
            System.Diagnostics.Debug.WriteLine($"[WordReader] 图片提取完成，共提取 {imagePaths.Count} 张图片");
            return imagePaths;
        }
    }
}

