using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Core;
using YamlDotNet.Core.Events;
using YamlDotNet.Serialization.TypeInspectors;
using YamlDotNet.Serialization.TypeResolvers;
using Newtonsoft.Json;
using WordAddIn1.DocumentHost;

namespace WordAddIn1
{
    /// <summary>
    /// 表格格式提取工具（只读）。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_ExtractTableFormatTool
    {
        /// <summary>
        /// YAML配置的数据结构
        /// </summary>
        public class TableConfig
        {
            public LocationInfo Location { get; set; }
            public TableInfo Table { get; set; }
            public List<List<int>> Merge { get; set; }
            public List<List<string>> Data { get; set; }
        }

        /// <summary>
        /// 位置信息
        /// </summary>
        public class LocationInfo
        {
            public string Type { get; set; }
            public int Start { get; set; }
            public int End { get; set; }
        }

        /// <summary>
        /// 表格信息
        /// </summary>
        public class TableInfo
        {
            public int Rows { get; set; }
            public int Cols { get; set; }
            public List<float> ColWidths { get; set; }
            public string Style { get; set; }
            public string TableFontColor { get; set; }
        }

        /// <summary>
        /// 注册表格格式提取工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_extract_table_format"] = async (args) =>
            {
                try
                {
                    // // System.Diagnostics.Debug.WriteLine("[DEBUG] extract_table_format工具开始执行");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError,
                            activateDocument: false))
                    {
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_extract_table_format] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    // 首先运行 WordReader.ReadWord 来生成表格编号和顺序序号映射表
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return new ToolResult { Success = false, Error = $"生成表格映射表失败：{ex.Message}" };
                    }

                    // 解析表格选择参数
                    var tableSelection = args["table_selection"] as Dictionary<string, object>;
                    var selection = TableExtractSelectionHelper.Resolve(document, tableSelection);
                    if (!selection.Success)
                    {
                        return new ToolResult { Success = false, Error = selection.Error };
                    }

                    Word.Table targetTable = selection.Table;
                    int tableIndex = selection.TableIndex;
                    string tableId = selection.TableId;
                    string selectionType = selection.SelectionType;

                    TableExtractDto dto;
                    try
                    {
                        dto = TableFormatExtractCore.Extract(targetTable, document);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[ExtractTableFormat] 提取失败：{ex.Message}");
                        return new ToolResult { Success = false, Error = $"表格格式提取失败: {ex.Message}" };
                    }

                    string filename = FilePathResolver.TryGetArg(args, "path")
                        ?? TableFormatFileHelper.GenerateFilename(
                        document,
                        tableIndex,
                        "format");
                    var (saved, xmlContent) = await SaveTableFormat(dto, filename);

                    return new ToolResult
                    {
                        Success = saved,
                        Error = saved ? null : "表格格式保存失败",
                        Data = new
                        {
                            table_extracted = saved,
                            table_info = new
                            {
                                rows = dto.Rows,
                                columns = dto.Cols,
                                table_index = tableIndex,
                                table_id = tableId,
                                selection_type = selectionType
                            },
                            format_file = filename,
                            format_saved = saved,
                            format_content = xmlContent,
                            warnings = dto.FontStats?.Warnings ?? new List<string>(),
                            message = saved
                                ? $"表格格式（General）提取并保存成功，文件名为：{filename}。单元格填空请使用 F_read_table_row_values。"
                                : "表格格式保存失败"
                        }
                    };
                }
                catch (Exception ex)
                {
                    // System.Diagnostics.Debug.WriteLine($"[DEBUG] extract_table_format工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"表格格式提取失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// 递归收集所有表格（包括嵌套表格），按照在文档中的位置（Range.Start）排序
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>按位置排序的所有表格列表</returns>
        private static List<Word.Table> GetAllTablesInOrder(Word.Document document)
        {
            List<Word.Table> allTables = new List<Word.Table>();
            
            try
            {
                // 收集所有顶级表格
                foreach (Word.Table table in document.Tables)
                {
                    CollectTablesRecursive(table, allTables);
                }
                
                // 按照 Range.Start 排序
                allTables.Sort((t1, t2) => t1.Range.Start.CompareTo(t2.Range.Start));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 收集表格失败：{ex.Message}");
            }
            
            return allTables;
        }

        /// <summary>
        /// 递归收集表格及其嵌套表格
        /// </summary>
        /// <param name="table">当前表格</param>
        /// <param name="allTables">收集到的表格列表</param>
        private static void CollectTablesRecursive(Word.Table table, List<Word.Table> allTables)
        {
            if (table == null)
            {
                return;
            }
            
            try
            {
                // 添加当前表格
                allTables.Add(table);
                
                // 遍历表格的所有单元格，查找嵌套表格（用 Range.Cells 避免纵向合并表 Rows 遍历失败）
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    foreach (Word.Table nestedTable in cell.Tables)
                    {
                        CollectTablesRecursive(nestedTable, allTables);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 递归收集表格失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 获取表格在文档中的序号
        /// </summary>
        private static int GetTableIndex(Word.Document document, Word.Table table)
        {
            try
            {
                for (int i = 1; i <= document.Tables.Count; i++)
                {
                    if (document.Tables[i].Range.Start == table.Range.Start)
                    {
                        return i;
                    }
                }
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取表格序号失败：{ex.Message}");
            }
            return 1; // 默认返回1
        }

        /// <summary>
        /// 提取表格格式（legacy 包装，供测试或内部复用）。
        /// </summary>
        private static TableConfig ExtractTableFormat(Word.Table table, int tableIndex, Word.Document document)
        {
            try
            {
                TableExtractDto dto = TableFormatExtractCore.Extract(table, document);
                return new TableConfig
                {
                    Location = new LocationInfo
                    {
                        Type = "char_pos",
                        Start = dto.RangeStart,
                        End = dto.RangeEnd,
                    },
                    Table = new TableInfo
                    {
                        Rows = dto.Rows,
                        Cols = dto.Cols,
                        ColWidths = dto.ColWidths,
                        Style = dto.Style,
                        TableFontColor = dto.FontStats?.TableFontColor ?? "000000",
                    },
                    Merge = dto.Merge,
                    Data = dto.Data,
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 表格格式提取失败：{ex.Message}");
                return null;
            }
        }

        // ExtractTableStyleFromXml 已移至 TableFormatExtractor

        /// <summary>
        /// 提取表格样式（保留原方法以兼容）
        /// </summary>
        private static string ExtractTableStyle(Word.Table table)
        {
            try
            {
                var style = table.get_Style();
                if (style != null)
                {
                    string styleName = style.ToString();
                    // System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取到表格样式：{styleName}");
                    return styleName;
                }
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取表格样式失败：{ex.Message}");
            }
            return "Grid Table 4 - Accent 1"; // 默认样式
        }

        // ExtractColWidthsFromXml 已移至 TableFormatExtractor


        /// <summary>
        /// 合并信息比较器，用于去重
        /// </summary>
        private class MergeInfoComparer : IEqualityComparer<List<int>>
        {
            public bool Equals(List<int> x, List<int> y)
            {
                if (x == null || y == null) return x == y;
                if (x.Count != y.Count) return false;
                for (int i = 0; i < x.Count; i++)
                {
                    if (x[i] != y[i]) return false;
                }
                return true;
            }

            public int GetHashCode(List<int> obj)
            {
                if (obj == null) return 0;
                int hash = 17;
                foreach (int value in obj)
                {
                    hash = hash * 31 + value.GetHashCode();
                }
                return hash;
            }
        }

        /// <summary>
        /// 整合连续的合并信息（行合并和列合并）
        /// </summary>
        private static List<List<int>> ConsolidateAllMerges(List<List<int>> merges)
        {
            var result = new List<List<int>>();

            // 按列分组，然后按行排序
            var groupedByColumn = merges
                .GroupBy(c => c[1]) // 按列号分组
                .ToDictionary(g => g.Key, g => g.OrderBy(c => c[0]).ToList()); // 按起始行排序

            foreach (var columnGroup in groupedByColumn)
            {
                int col = columnGroup.Key;
                var columnMerges = columnGroup.Value;

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 整合列{col}的合并，共有{columnMerges.Count}个记录");

                // 合并连续的行范围
                int startRow = -1;
                int endRow = -1;

                foreach (var merge in columnMerges)
                {
                    int currentStartRow = merge[0];
                    int currentEndRow = merge[2];

                    if (startRow == -1)
                    {
                        // 第一个合并
                        startRow = currentStartRow;
                        endRow = currentEndRow;
                    }
                    else if (currentStartRow == endRow)
                    {
                        // 连续的合并，扩展结束行
                        endRow = currentEndRow;
                    }
                    else
                    {
                        // 不连续，保存之前的合并，开始新的合并
                        result.Add(new List<int> { startRow, col, endRow, col });
                        System.Diagnostics.Debug.WriteLine($"[DEBUG]   保存行合并: [{startRow}, {col}, {endRow}, {col}]");

                        startRow = currentStartRow;
                        endRow = currentEndRow;
                    }
                }

                // 保存最后一个合并
                if (startRow != -1)
                {
                    result.Add(new List<int> { startRow, col, endRow, col });
                    System.Diagnostics.Debug.WriteLine($"[DEBUG]   保存行合并: [{startRow}, {col}, {endRow}, {col}]");
                }
            }

            // 保留列合并（行相等，列不同的合并）
            var existingColumnMerges = merges.Where(m => m[0] == m[2] && m[1] != m[3]).ToList();
            result.AddRange(existingColumnMerges);

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 整合完成，得到{result.Count}个合并范围（包含行合并和列合并）");
            return result;
        }

        /// <summary>
        /// 整理行合并：将首尾相接的行合并记录合并
        /// 例如：[[2,3,3,3],[3,3,4,3]] 合并为 [[2,3,4,3]]
        /// </summary>
        private static List<List<int>> ConsolidateRowMerges(List<List<int>> merges)
        {
            var result = new List<List<int>>();

            // 分离出行合并和列合并
            var rowMerges = merges.Where(m => m[1] == m[3]).ToList(); // 行合并：startCol == endCol
            var columnMerges = merges.Where(m => m[0] == m[2] && m[1] != m[3]).ToList(); // 列合并：startRow == endRow 且 startCol != endCol

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 整理前：行合并 {rowMerges.Count} 个，列合并 {columnMerges.Count} 个");

            // 按列分组处理行合并
            var groupedByColumn = rowMerges
                .GroupBy(m => m[1]) // 按列号分组
                .ToDictionary(g => g.Key, g => g.OrderBy(m => m[0]).ToList()); // 按起始行排序

            foreach (var columnGroup in groupedByColumn)
            {
                int col = columnGroup.Key;
                var mergesForColumn = columnGroup.Value;

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 整理列{col}的行合并，共有{mergesForColumn.Count}个记录");

                // 合并首尾相接的记录
                int startRow = -1;
                int endRow = -1;

                foreach (var merge in mergesForColumn)
                {
                    int currentStartRow = merge[0];
                    int currentEndRow = merge[2];

                    if (startRow == -1)
                    {
                        // 第一个合并
                        startRow = currentStartRow;
                        endRow = currentEndRow;
                    }
                    else if (currentStartRow == endRow + 1 || currentStartRow == endRow)
                    {
                        // 首尾相接（endRow + 1）或重叠（endRow），扩展结束行
                        endRow = Math.Max(endRow, currentEndRow);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG]   合并记录: [{startRow}, {col}, {endRow}, {col}]");
                    }
                    else
                    {
                        // 不连续，保存之前的合并，开始新的合并
                        result.Add(new List<int> { startRow, col, endRow, col });
                        System.Diagnostics.Debug.WriteLine($"[DEBUG]   保存行合并: [{startRow}, {col}, {endRow}, {col}]");

                        startRow = currentStartRow;
                        endRow = currentEndRow;
                    }
                }

                // 保存最后一个合并
                if (startRow != -1)
                {
                    result.Add(new List<int> { startRow, col, endRow, col });
                    System.Diagnostics.Debug.WriteLine($"[DEBUG]   保存行合并: [{startRow}, {col}, {endRow}, {col}]");
                }
            }

            // 添加列合并
            result.AddRange(columnMerges);

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 整理后：共 {result.Count} 个合并记录（行合并 {result.Count - columnMerges.Count} 个，列合并 {columnMerges.Count} 个）");
            return result;
        }

        /// <summary>
        /// 打印 ColumnMergeInfo（列对应关系）
        /// </summary>
        private static void PrintColumnMergeInfo(Dictionary<int, List<List<int>>> columnMapping, string context = "")
        {
            if (!string.IsNullOrEmpty(context))
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ========== ColumnMergeInfo {context} ==========");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ========== ColumnMergeInfo ==========");
            }
            
            var result = new List<object>();
            foreach (var kvp in columnMapping.OrderBy(x => x.Key))
            {
                int row = kvp.Key;
                var mapping = kvp.Value;
                result.Add(new List<object> { row, mapping });
            }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < result.Count; i++)
            {
                var rowMapping = result[i] as List<object>;
                int row = (int)rowMapping[0];
                var mapping = rowMapping[1] as List<List<int>>;
                
                sb.Append($"  [{row},[");
                for (int j = 0; j < mapping.Count; j++)
                {
                    sb.Append($"[{mapping[j][0]},{mapping[j][1]}]");
                    if (j < mapping.Count - 1)
                        sb.Append(",");
                }
                sb.Append("]]");
                if (i < result.Count - 1)
                    sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");

            string mappingStr = sb.ToString();
            System.Diagnostics.Debug.WriteLine(mappingStr);
            System.Diagnostics.Debug.WriteLine($"[DEBUG] =====================================");
        }

        /// <summary>
        /// 打印 RowMergeInfo（行合并记录，只包含行合并，不包含列合并）
        /// </summary>
        private static void PrintRowMergeInfo(List<List<int>> merge, string context = "")
        {
            if (!string.IsNullOrEmpty(context))
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ========== RowMergeInfo {context} ==========");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ========== RowMergeInfo ==========");
            }
            
            // 在第二步时，不进行过滤，直接打印所有记录
            // 在其他步骤时，只过滤出行合并记录（startCol == endCol）
            var rowMerges = context.Contains("第二步") ? merge : merge.Where(m => m[1] == m[3]).ToList();
            
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("[");
            for (int i = 0; i < rowMerges.Count; i++)
            {
                var m = rowMerges[i];
                sb.Append($"  [{m[0]},{m[1]},{m[2]},{m[3]}]");
                if (i < rowMerges.Count - 1)
                    sb.Append(",");
                sb.AppendLine();
            }
            sb.AppendLine("]");
            
            string mergeStr = sb.ToString();
            System.Diagnostics.Debug.WriteLine(mergeStr);
            System.Diagnostics.Debug.WriteLine($"[DEBUG] =====================================");
        }

        // ExtractMergeInfoAndDataFromXml, ProcessCellFromXml, ExtractCellContentFromXml, 
        // GetValAttributeFromXml, CalculateAllRowspansFromXml 已移至 TableFormatExtractor



        /// <summary>
        /// 统计表格中所有单元格的颜色频率
        /// </summary>
        private static string CalculateMostCommonColor(Word.Table table)
        {
            var colorCounts = new Dictionary<string, int>();
            int totalRuns = 0;

            try
            {
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        try
                        {
                            Word.Cell cell = table.Cell(row, col);
                            string xml = cell.Range.WordOpenXML;
                            XDocument doc = XDocument.Parse(xml);
                            XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                            var runs = doc.Descendants(w + "r");
                            foreach (var r in runs)
                            {
                                totalRuns++;
                                var rPr = r.Element(w + "rPr");
                                string color = rPr?.Elements(w + "color").FirstOrDefault()?.Attribute(w + "val")?.Value;

                                // 如果没有显式颜色，认为是自动颜色（通常是黑色或继承颜色）
                                string effectiveColor;
                                if (string.IsNullOrEmpty(color))
                                {
                                    effectiveColor = "auto";
                                }
                                else
                                {
                                    // Word使用BGR格式，需要转换为RGB格式
                                    effectiveColor = TableFormatExtractor.ConvertBgrToRgb(color.ToUpper());
                                }

                                if (!colorCounts.ContainsKey(effectiveColor))
                                    colorCounts[effectiveColor] = 0;
                                colorCounts[effectiveColor]++;
                            }
                        }
                        catch (Exception)
                        {
                            // System.Diagnostics.Debug.WriteLine($"[DEBUG] 统计单元格({row},{col})颜色失败：{ex.Message}");
                        }
                    }
                }

                // 找出出现次数最多的颜色（排除auto，因为auto表示默认颜色）
                var explicitColors = colorCounts.Where(kvp => kvp.Key != "auto");
                if (explicitColors.Any())
                {
                    var mostCommonColor = explicitColors.OrderByDescending(kvp => kvp.Value).First();
                    // System.Diagnostics.Debug.WriteLine($"[DEBUG] 最常见的显式颜色：#{mostCommonColor.Key}，出现次数：{mostCommonColor.Value}，总文本段数：{totalRuns}");
                    return mostCommonColor.Key;
                }
                else
                {
                    // System.Diagnostics.Debug.WriteLine($"[DEBUG] 没有找到显式颜色设置，默认使用黑色，总文本段数：{totalRuns}");
                }
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 统计颜色失败：{ex.Message}");
            }

            return "000000"; // 默认黑色
        }

        // ConvertBgrToRgb 已移至 TableFormatExtractor

        /// <summary>
        /// 将单元格格式转换为HTML
        /// </summary>
        private static string CellToHtml(Word.Cell cell, string defaultColor = null)
        {
            try
            {
                // 取出单元格的 WordOpenXML
                string xml = cell.Range.WordOpenXML;
                XDocument doc = XDocument.Parse(xml);
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

                var runs = doc.Descendants(w + "r");
                if (!runs.Any()) return cell.Range.Text;

                var sb = new StringBuilder();
                foreach (var r in runs)
                {
                    string txt = string.Join("", r.Descendants(w + "t").Select(t => (string)t));
                    if (string.IsNullOrEmpty(txt)) continue;

                    var rPr = r.Element(w + "rPr");
                    bool bold = rPr?.Elements(w + "b").Any() == true;
                    bool italic = rPr?.Elements(w + "i").Any() == true;
                    string color = rPr?.Elements(w + "color").FirstOrDefault()?.Attribute(w + "val")?.Value;

                    // 如果有默认颜色，且当前颜色与默认颜色相同，或者是auto颜色，则不添加颜色样式
                    bool shouldIncludeColor = !string.IsNullOrEmpty(color) &&
                                            color.ToUpper() != "auto" &&
                                            (string.IsNullOrEmpty(defaultColor) || !TableFormatExtractor.ConvertBgrToRgb(color).Equals(defaultColor, StringComparison.OrdinalIgnoreCase));

                    // 拼最小 HTML
                    if (bold || italic || shouldIncludeColor)
                    {
                        sb.Append("<span");
                        var styles = new List<string>();
                        if (bold) styles.Add("font-weight:bold");
                        if (italic) styles.Add("font-style:italic");
                        if (shouldIncludeColor) styles.Add($"color:#{TableFormatExtractor.ConvertBgrToRgb(color)}");

                        if (styles.Any())
                        {
                            sb.Append($" style='{string.Join(";", styles)}'");
                        }
                        sb.Append($">{System.Web.HttpUtility.HtmlEncode(txt)}</span>");
                    }
                    else
                    {
                        sb.Append(System.Web.HttpUtility.HtmlEncode(txt));
                    }
                }
                return sb.ToString();
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 单元格HTML转换失败：{ex.Message}");
                // 降级到纯文本
                return cell.Range.Text?.Trim() ?? "";
            }
        }

        /// <summary>
        /// 检查指定位置是否在合并范围内（被合并覆盖）
        /// </summary>
        private static bool IsInMergeRange(int row, int col, List<List<int>> merge)
        {
            foreach (var mergeInfo in merge)
            {
                // mergeInfo格式: [startRow, startCol, endRow, endCol]
                int startRow = mergeInfo[0], startCol = mergeInfo[1];
                int endRow = mergeInfo[2], endCol = mergeInfo[3];

                if (row >= startRow && row <= endRow && col >= startCol && col <= endCol)
                {
                    // 如果是起始点，返回false（起始点有数据）
                    if (row == startRow && col == startCol)
                        return false;
                    // 其他点都被覆盖，返回true
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 提取表格数据
        /// </summary>
        private static void ExtractTableData(Word.Table table, List<List<string>> data, string defaultColor, List<List<int>> merge)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 开始提取表格数据，基于{merge.Count}个合并信息进行重新排列");

                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    // 1. 先用原来的逻辑提取Word实际存在的单元格数据
                    var rawRowData = new List<string>();
                    int actualColIndex = 0;

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取第{row}行的原始数据...");

                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        try
                        {
                            Word.Cell cell = table.Cell(row, col);

                            // 检查这个单元格是否是合并单元格的左上角
                            if (cell.RowIndex == row && cell.ColumnIndex == col)
                            {
                                // 这是一个有效的单元格，提取内容
                                string cellHtml = CellToHtml(cell, defaultColor);

                                if (string.IsNullOrWhiteSpace(cellHtml))
                                {
                                    string cellText = cell.Range.Text?.Trim() ?? "";
                                    if (cellText.EndsWith("\r\a"))
                                        cellText = cellText.Substring(0, cellText.Length - 2);
                                    else if (cellText.EndsWith("\r"))
                                        cellText = cellText.Substring(0, cellText.Length - 1);

                                    rawRowData.Add(cellText);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col} (实际索引{actualColIndex + 1}): '{cellText}'");
                                }
                                else
                                {
                                    rawRowData.Add(cellHtml);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col} (实际索引{actualColIndex + 1}): '{cellHtml}'");
                                }

                                actualColIndex++;
                            }
                            // 其他情况（被合并或不存在）跳过，由后续逻辑处理
                        }
                        catch (Exception)
                        {
                            // 这个位置不存在单元格
                            System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col}: 不存在");
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{row}行原始数据: {rawRowData.Count}个有效单元格");

                    // 2. 根据Merge信息重新排列数据
                    var finalRowData = new List<string>();
                    int rawDataIndex = 0;

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 根据合并信息重新排列第{row}行数据...");

                    for (int col = 1; col <= table.Columns.Count; col++)
                    {
                        if (IsInMergeRange(row, col, merge))
                        {
                            // 这个位置被合并覆盖，添加空字符串
                            finalRowData.Add("");
                            System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col}: 空（被合并覆盖）");
                        }
                        else
                        {
                            // 这个位置应该有数据
                            if (rawDataIndex < rawRowData.Count)
                            {
                                finalRowData.Add(rawRowData[rawDataIndex]);
                                System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col}: '{rawRowData[rawDataIndex]}'（来自原始数据索引{rawDataIndex}）");
                                rawDataIndex++;
                            }
                            else
                            {
                                // 不应该发生的情况，但为了安全
                                finalRowData.Add("");
                                System.Diagnostics.Debug.WriteLine($"[DEBUG]   列{col}: 空（无原始数据）");
                            }
                        }
                    }

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 第{row}行最终数据: [{string.Join(", ", finalRowData.Select(s => $"\"{s}\""))}]");
                    data.Add(finalRowData);
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 表格数据提取完成：{data.Count}行，每行{table.Columns.Count}列");
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取表格数据失败：{ex.Message}");
            }
        }

        /// <summary>
        /// 生成文件名
        /// </summary>
        private static string GenerateFilename(Word.Document document, int tableIndex)
        {
            try
            {
                // 获取文档文件名（不含扩展名）
                string docName = Path.GetFileNameWithoutExtension(document.Name);

                // 获取当前时间戳
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");

                // 生成文件名：文档名_第几个表格_format_时间戳.xml
                string filename = $"{docName}_{tableIndex}_format_{timestamp}.xml";

                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 生成文件名：{filename}");
                return filename;
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 生成文件名失败：{ex.Message}");
                // 使用默认文件名
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                return $"table_{tableIndex}_format_{timestamp}.xml";
            }
        }

        /// <summary>
        /// 保存表格格式到文件（仅 General）。
        /// </summary>
        private static async Task<(bool Success, string XmlContent)> SaveTableFormat(
            TableExtractDto dto,
            string filename)
        {
            try
            {
                string xmlContent = TableFormatXmlBuilder.BuildGeneralOnly(dto);
                return await TableFormatFileHelper.SaveXmlAsync(filename, xmlContent);
            }
            catch (Exception)
            {
                return (false, null);
            }
        }

        /// <summary>
        /// 将数值数组转换为流式格式，并修复Data字段格式
        /// </summary>
        private static string ConvertNumericArraysToFlowStyle(string yamlContent)
        {
            var lines = yamlContent.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var result = new List<string>();
            var inDataSection = false;
            var dataRows = new List<List<string>>();

            for (int i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                var trimmed = line.Trim();

                // 检查是否进入Data字段
                if (trimmed.StartsWith("Data:"))
                {
                    inDataSection = true;
                    result.Add(line);
                    continue;
                }

                if (inDataSection)
                {
                    // 检查是否离开Data字段
                    if (!trimmed.StartsWith(" ") && !trimmed.StartsWith("-") && trimmed.Contains(":") && !trimmed.StartsWith("Data:"))
                    {
                        // 处理完Data字段，转换为正确格式
                        if (dataRows.Count > 0)
                        {
                            foreach (var row in dataRows)
                            {
                                var rowArray = "[" + string.Join(", ", row.Select(s => "'" + s.Replace("'", "''") + "'")) + "]";
                                result.Add("  - " + rowArray);
                            }
                        }
                        inDataSection = false;
                        result.Add(line); // 添加下一个字段
                        continue;
                    }

                    // 解析Data字段的内容
                    if (trimmed.StartsWith("- "))
                    {
                        var rowData = new List<string>();
                        var rowIndent = line.Length - trimmed.Length;

                        // 收集这一行的所有列
                        rowData.Add(trimmed.Substring(2));

                        // 查找同一行的其他列
                        var k = i + 1;
                        while (k < lines.Length)
                        {
                            var nextLine = lines[k];
                            var nextTrimmed = nextLine.Trim();
                            var nextIndent = nextLine.Length - nextTrimmed.Length;

                            if (nextIndent > rowIndent && nextTrimmed.StartsWith("- "))
                            {
                                rowData.Add(nextTrimmed.Substring(2));
                                k++;
                            }
                            else
                            {
                                break;
                            }
                        }

                        dataRows.Add(rowData);
                        i = k - 1; // 跳过已处理的行
                    }
                }
                else
                {
                    // 检查是否是数值数组字段（以冒号结尾的字段名后跟数值列表）
                    if (trimmed.EndsWith(":") && !trimmed.StartsWith("Data:") && !trimmed.StartsWith("Merge:"))
                    {
                        var fieldName = trimmed.TrimEnd(':');
                        var fieldIndent = line.Length - trimmed.Length;

                        // 检查接下来的行是否是数值数组
                        var arrayLines = new List<string>();
                        var j = i + 1;

                        while (j < lines.Length)
                        {
                            var nextLine = lines[j];
                            var nextTrimmed = nextLine.Trim();

                            if (string.IsNullOrWhiteSpace(nextLine))
                            {
                                // 空行，跳过
                                j++;
                                continue;
                            }

                            var nextIndent = nextLine.Length - nextTrimmed.Length;

                            // 如果缩进小于或等于字段缩进，说明这个字段结束了
                            if (nextIndent <= fieldIndent && nextTrimmed.Contains(":"))
                            {
                                break;
                            }

                            // 如果是数值行（以 - 开头且后面是数字）
                            if (nextTrimmed.StartsWith("- ") && IsNumericValue(nextTrimmed.Substring(2).Trim()))
                            {
                                arrayLines.Add(nextTrimmed.Substring(2).Trim());
                                j++;
                            }
                            else
                            {
                                // 不是数值行，跳出
                                break;
                            }
                        }

                        // 如果找到至少2个数值，转换为流式格式
                        if (arrayLines.Count >= 2)
                        {
                            // 直接在字段名后添加数组，避免额外缩进
                            var flowSequence = "[" + string.Join(", ", arrayLines) + "]";
                            result.Add(line.Replace(":", ": " + flowSequence));

                            // 跳过已经处理的数组行
                            i = j - 1;
                        }
                        else
                        {
                            // 保持原样
                            result.Add(line);
                        }
                    }
                    else
                    {
                        result.Add(line);
                    }
                }
            }

            // 处理文件末尾的Data字段
            if (inDataSection && dataRows.Count > 0)
            {
                foreach (var row in dataRows)
                {
                    var rowArray = "[" + string.Join(", ", row.Select(s => "'" + s.Replace("'", "''") + "'")) + "]";
                    result.Add("  - " + rowArray);
                }
            }

            return string.Join("\n", result);
        }

        /// <summary>
        /// 检查字符串是否是数值
        /// </summary>
        private static bool IsNumericValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            // 检查是否是数字（包括浮点数）
            return double.TryParse(value, out _);
        }

        /// <summary>
        /// 自定义类型转换器：确保List<float>序列化为流式格式
        /// </summary>
        private class ListFloatConverter : IYamlTypeConverter
        {
            public bool Accepts(Type type)
            {
                return type == typeof(List<float>);
            }

            public object ReadYaml(IParser parser, Type type)
            {
                // 反序列化时保持默认行为
                var deserializer = new DeserializerBuilder().Build();
                return deserializer.Deserialize<List<float>>(parser);
            }

            public void WriteYaml(IEmitter emitter, object value, Type type)
            {
                var list = (List<float>)value;
                emitter.Emit(new SequenceStart(null, null, true, SequenceStyle.Flow));

                foreach (var item in list)
                {
                    emitter.Emit(new Scalar(item.ToString()));
                }

                emitter.Emit(new SequenceEnd());
            }
        }

        /// <summary>
        /// 自定义类型转换器：确保List<List<string>>序列化为流式格式
        /// </summary>
        private class ListListStringConverter : IYamlTypeConverter
        {
            public bool Accepts(Type type)
            {
                return type == typeof(List<List<string>>);
            }

            public object ReadYaml(IParser parser, Type type)
            {
                // 反序列化时保持默认行为
                var deserializer = new DeserializerBuilder().Build();
                return deserializer.Deserialize<List<List<string>>>(parser);
            }

            public void WriteYaml(IEmitter emitter, object value, Type type)
            {
                var list = (List<List<string>>)value;
                emitter.Emit(new SequenceStart(null, null, false, SequenceStyle.Flow)); // 使用流式序列

                foreach (var subList in list)
                {
                    emitter.Emit(new SequenceStart(null, null, true, SequenceStyle.Flow)); // 每个子数组使用流式
                    foreach (var item in subList)
                    {
                        emitter.Emit(new Scalar(item ?? ""));
                    }
                    emitter.Emit(new SequenceEnd());
                }

                emitter.Emit(new SequenceEnd());
            }
        }
    }
}
