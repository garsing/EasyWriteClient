using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 表格格式提取工具
    /// 提供从Word表格中提取格式信息的各种方法
    /// </summary>
    public static class TableFormatExtractor
    {
        /// <summary>
        /// 单元格信息
        /// </summary>
        public class CellInfo
        {
            public int StartRow { get; set; }
            public int StartCol { get; set; }
            public int EndRow { get; set; }
            public int EndCol { get; set; }
            public int Colspan { get; set; }
            public int Rowspan { get; set; }
            public bool HasMerge { get; set; }
            public bool IsVerticalMergeContinue { get; set; }
            public string Content { get; set; }
        }

        /// <summary>
        /// 获取元素的 val 属性值
        /// </summary>
        public static string GetValAttributeFromXml(XElement element)
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
        /// 将 BGR 颜色格式转换为 RGB 格式（Word OOXML/COM 内部为 BGR，配置 XML 存 RGB）。
        /// 例如: 0000FF(BGR 红) -> FF0000(RGB 红)
        /// </summary>
        public static string ConvertBgrToRgb(string bgrColor)
        {
            try
            {
                if (bgrColor.Length != 6)
                    return bgrColor;

                string b = bgrColor.Substring(0, 2);
                string g = bgrColor.Substring(2, 2);
                string r = bgrColor.Substring(4, 2);

                return r + g + b;
            }
            catch
            {
                return bgrColor;
            }
        }

        /// <summary>
        /// 将 RGB 颜色格式转换为 BGR 格式（写入 Word OOXML/COM 时使用）。
        /// 例如: FF0000(RGB 红) -> 0000FF(BGR 红)
        /// </summary>
        public static string ConvertRgbToBgr(string rgbColor)
        {
            try
            {
                if (rgbColor.Length != 6)
                    return rgbColor;

                string r = rgbColor.Substring(0, 2);
                string g = rgbColor.Substring(2, 2);
                string b = rgbColor.Substring(4, 2);

                return b + g + r;
            }
            catch
            {
                return rgbColor;
            }
        }

        /// <summary>
        /// 配置 XML 中的 RGB 十六进制 → Word COM WdColor（BGR 整型）。
        /// </summary>
        public static Word.WdColor ConvertRgbHexToWdColor(string rgbHex)
        {
            string hex = (rgbHex ?? string.Empty).Trim().TrimStart('#');
            if (hex.Length != 6)
            {
                return Word.WdColor.wdColorAutomatic;
            }

            string bgrHex = ConvertRgbToBgr(hex.ToUpperInvariant());
            return (Word.WdColor)Convert.ToInt32(bgrHex, 16);
        }

        /// <summary>
        /// 预先扫描表格，计算所有纵向合并的rowspan值
        /// </summary>
        public static Dictionary<(int row, int col), int> CalculateAllRowspansFromXml(List<XElement> rows, XNamespace w)
        {
            var rowspanMap = new Dictionary<(int row, int col), int>();
            var verticalMergeStarts = new Dictionary<int, (int startRow, int startCol, int gridSpan)>();
            
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
                        string gridSpanVal = GetValAttributeFromXml(gridSpanElem);
                        if (!string.IsNullOrEmpty(gridSpanVal))
                        {
                            int.TryParse(gridSpanVal, out gridSpan);
                        }
                    }
                    
                    // 检查vMerge
                    var vMerge = tcPr?.Element(w + "vMerge");
                    if (vMerge != null)
                    {
                        string vMergeVal = GetValAttributeFromXml(vMerge);
                        if (vMergeVal == "restart")
                        {
                            // 记录纵向合并起始
                            verticalMergeStarts[logicalColIdx] = (rowIdx, logicalColIdx, gridSpan);
                        }
                    }
                    
                    logicalColIdx += gridSpan;
                }
            }
            
            // 第二遍扫描：计算每个纵向合并的rowspan
            foreach (var kvp in verticalMergeStarts)
            {
                int startCol = kvp.Key;
                int startRow = kvp.Value.startRow;
                int gridSpan = kvp.Value.gridSpan;
                
                // 从起始行开始，向下查找延续单元格
                int rowspan = 1;
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
                            string gridSpanVal = GetValAttributeFromXml(gridSpanElem);
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
                                string vMergeVal = GetValAttributeFromXml(vMerge);
                                if (vMergeVal != "restart")  // 延续单元格：val 为空或不是"restart"
                                {
                                    // 找到延续单元格
                                    foundContinue = true;
                                    rowspan++;
                                    break;
                                }
                            }
                        }
                        
                        logicalColIdx += cellGridSpan;
                    }
                    
                    if (!foundContinue)
                    {
                        // 没有找到延续单元格，合并结束
                        break;
                    }
                }
                
                // 记录rowspan值（起始单元格和所有延续单元格都使用相同的rowspan）
                for (int r = startRow; r < startRow + rowspan; r++)
                {
                    rowspanMap[(r, startCol)] = rowspan;
                }
            }
            
            return rowspanMap;
        }

        /// <summary>
        /// 提取表格样式（从XML）
        /// </summary>
        public static string ExtractTableStyleFromXml(XElement tableElement, XNamespace w, Word.Table table)
        {
            try
            {
                // 首先尝试从 XML 中提取样式
                var tblPr = tableElement.Element(w + "tblPr");
                if (tblPr != null)
                {
                    var tblStyle = tblPr.Element(w + "tblStyle");
                    if (tblStyle != null)
                    {
                        var valAttr = tblStyle.Attribute(w + "val");
                        if (valAttr != null && !string.IsNullOrEmpty(valAttr.Value))
                        {
                            return valAttr.Value;
                        }
                    }
                }

                if (table != null)
                {
                    var style = table.get_Style();
                    if (style != null)
                    {
                        return style.ToString();
                    }
                }
            }
            catch (Exception)
            {
                // System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取表格样式失败：{ex.Message}");
            }
            return "Grid Table 4 - Accent 1"; // 默认样式
        }

        /// <summary>
        /// 从XML提取列宽
        /// </summary>
        public static List<float> ExtractColWidthsFromXml(XElement tableElement, XNamespace w, int cols)
        {
            var colWidths = new List<float>();
            
            try
            {
                // 从表格属性中获取列定义
                var tblGrid = tableElement.Element(w + "tblGrid");
                if (tblGrid != null)
                {
                    var gridCols = tblGrid.Elements(w + "gridCol").ToList();
                    foreach (var gridCol in gridCols)
                    {
                        var wAttr = gridCol.Attribute(w + "w");
                        if (wAttr != null && int.TryParse(wAttr.Value, out int width))
                        {
                            // Word XML 中的宽度单位是 twips (1/20 point)，需要转换为 points
                            colWidths.Add(width / 20.0f);
                        }
                        else
                        {
                            colWidths.Add(80.0f); // 默认值
                        }
                    }
                }

                // 如果从 XML 中提取的列数不够，补充默认值
                while (colWidths.Count < cols)
                {
                    colWidths.Add(80.0f);
                }

                // 如果提取的列数超过需要的，截断
                if (colWidths.Count > cols)
                {
                    colWidths = colWidths.Take(cols).ToList();
                }
            }
            catch (Exception)
            {
                // 如果提取失败，使用默认值
                for (int i = 0; i < cols; i++)
                {
                    colWidths.Add(80.0f);
                }
            }

            return colWidths;
        }

        /// <summary>
        /// 从XML提取单元格内容
        /// </summary>
        public static string ExtractCellContentFromXml(
            XElement cell,
            XNamespace w,
            Dictionary<string, int> colorCounts,
            ref int totalRuns)
        {
            var sb = new StringBuilder();
            
            try
            {
                // 提取段落
                var paragraphs = cell.Elements(w + "p");
                foreach (var para in paragraphs)
                {
                    // 提取文本运行
                    var runs = para.Elements(w + "r");
                    foreach (var r in runs)
                    {
                        string txt = string.Join("", r.Descendants(w + "t").Select(t => (string)t));
                        if (string.IsNullOrEmpty(txt)) continue;

                        totalRuns++;
                        var rPr = r.Element(w + "rPr");
                        
                        // 统计颜色
                        string color = rPr?.Elements(w + "color").FirstOrDefault()?.Attribute(w + "val")?.Value;
                        string effectiveColor;
                        if (string.IsNullOrEmpty(color))
                        {
                            effectiveColor = "auto";
                        }
                        else
                        {
                            effectiveColor = ConvertBgrToRgb(color.ToUpper());
                        }

                        if (!colorCounts.ContainsKey(effectiveColor))
                            colorCounts[effectiveColor] = 0;
                        colorCounts[effectiveColor]++;

                        // 提取格式信息
                        bool bold = rPr?.Elements(w + "b").Any() == true;
                        bool italic = rPr?.Elements(w + "i").Any() == true;

                        // 构建HTML内容
                        if (bold || italic || (!string.IsNullOrEmpty(color) && color.ToUpper() != "AUTO"))
                        {
                            sb.Append("<span");
                            var styles = new List<string>();
                            if (bold) styles.Add("font-weight:bold");
                            if (italic) styles.Add("font-style:italic");
                            if (!string.IsNullOrEmpty(color) && color.ToUpper() != "AUTO")
                            {
                                styles.Add($"color:#{effectiveColor}");
                            }

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
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取单元格内容失败：{ex.Message}");
            }

            string content = sb.ToString();
            // 清理末尾的换行符
            if (content.EndsWith("\r\a"))
                content = content.Substring(0, content.Length - 2);
            else if (content.EndsWith("\r"))
                content = content.Substring(0, content.Length - 1);

            return content;
        }

        /// <summary>
        /// 处理单元格，提取合并信息和内容
        /// </summary>
        public static CellInfo ProcessCellFromXml(
            XElement cell,
            XNamespace w,
            int currentRowIndex,
            int currentLogicalColIndex,
            Dictionary<(int row, int col), int> rowspanMap,
            Dictionary<string, int> colorCounts,
            ref int totalRuns)
        {
            var cellInfo = new CellInfo
            {
                StartRow = currentRowIndex,
                StartCol = currentLogicalColIndex,
                EndRow = currentRowIndex,
                EndCol = currentLogicalColIndex,
                Colspan = 1,
                Rowspan = 1,
                HasMerge = false,
                IsVerticalMergeContinue = false,
                Content = ""
            };

            try
            {
                var tcPr = cell.Element(w + "tcPr");

                // 1. 提取 gridSpan（横向合并）
                var gridSpan = tcPr?.Element(w + "gridSpan");
                if (gridSpan != null)
                {
                    string gridSpanVal = GetValAttributeFromXml(gridSpan);
                    if (!string.IsNullOrEmpty(gridSpanVal) && int.TryParse(gridSpanVal, out int parsedColspan))
                    {
                        cellInfo.Colspan = parsedColspan;
                    }
                }

                // 2. 提取 vMerge（纵向合并）
                var vMerge = tcPr?.Element(w + "vMerge");
                if (vMerge != null)
                {
                    string vMergeVal = GetValAttributeFromXml(vMerge);
                    if (vMergeVal == "restart")
                    {
                        // 纵向合并起始
                        if (rowspanMap.TryGetValue((currentRowIndex, currentLogicalColIndex), out int calculatedRowspan))
                        {
                            cellInfo.Rowspan = calculatedRowspan;
                        }
                    }
                    else
                    {
                        // 纵向合并延续
                        cellInfo.IsVerticalMergeContinue = true;
                        if (rowspanMap.TryGetValue((currentRowIndex, currentLogicalColIndex), out int calculatedRowspan))
                        {
                            cellInfo.Rowspan = calculatedRowspan;
                        }
                        // 延续单元格不输出内容，直接返回
                        return cellInfo;
                    }
                }

                // 3. 计算合并范围
                cellInfo.EndRow = cellInfo.StartRow + cellInfo.Rowspan - 1;
                cellInfo.EndCol = cellInfo.StartCol + cellInfo.Colspan - 1;
                cellInfo.HasMerge = (cellInfo.Colspan > 1) || (cellInfo.Rowspan > 1);

                // 4. 提取单元格内容
                cellInfo.Content = ExtractCellContentFromXml(cell, w, colorCounts, ref totalRuns);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 处理单元格失败：{ex.Message}");
            }

            return cellInfo;
        }

        /// <summary>
        /// 从XML提取合并信息和数据
        /// </summary>
        public static string ExtractMergeInfoAndDataFromXml(
            XElement tableElement, 
            XNamespace w, 
            int rows, 
            int cols, 
            List<List<int>> merge, 
            List<List<string>> data)
        {
            var colorCounts = new Dictionary<string, int>();
            int totalRuns = 0;

            try
            {
                // 获取所有行
                var xmlRows = tableElement.Elements(w + "tr").ToList();
                
                // 预先扫描表格，计算所有纵向合并的 rowspan
                var rowspanMap = CalculateAllRowspansFromXml(xmlRows, w);

                // 初始化数据矩阵（rows x cols）
                for (int r = 0; r < rows; r++)
                {
                    data.Add(new List<string>());
                    for (int c = 0; c < cols; c++)
                    {
                        data[r].Add(""); // 初始化为空字符串
                    }
                }

                // 处理每一行
                int currentRowIndex = 0; // 0-based
                int currentLogicalColIndex = 0; // 逻辑列索引（考虑gridSpan）

                foreach (var row in xmlRows)
                {
                    if (currentRowIndex >= rows) break;

                    currentLogicalColIndex = 0;
                    var cells = row.Elements(w + "tc").ToList();

                    foreach (var cell in cells)
                    {
                        // 提取单元格合并信息和内容
                        var cellInfo = ProcessCellFromXml(
                            cell, 
                            w, 
                            currentRowIndex, 
                            currentLogicalColIndex, 
                            rowspanMap,
                            colorCounts,
                            ref totalRuns);

                        // 如果有合并信息，添加到 merge 列表
                        if (cellInfo.HasMerge)
                        {
                            // 格式：[起始行, 起始列, 结束行, 结束列] (1-based)
                            merge.Add(new List<int>
                            {
                                cellInfo.StartRow + 1,      // 转换为1-based
                                cellInfo.StartCol + 1,      // 转换为1-based
                                cellInfo.EndRow + 1,        // 转换为1-based
                                cellInfo.EndCol + 1         // 转换为1-based
                            });
                        }

                        // 如果是合并的起始单元格，保存数据
                        if (!cellInfo.IsVerticalMergeContinue)
                        {
                            int dataRow = cellInfo.StartRow;
                            int dataCol = cellInfo.StartCol;
                            
                            if (dataRow < rows && dataCol < cols)
                            {
                                data[dataRow][dataCol] = cellInfo.Content;
                            }
                        }

                        // 更新逻辑列索引
                        currentLogicalColIndex += cellInfo.Colspan;
                    }

                    currentRowIndex++;
                }

                // 计算最常见的颜色
                var explicitColors = colorCounts.Where(kvp => kvp.Key != "auto");
                if (explicitColors.Any())
                {
                    var mostCommonColor = explicitColors.OrderByDescending(kvp => kvp.Value).First();
                    return mostCommonColor.Key;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 提取合并信息和数据失败：{ex.Message}");
            }

            return "000000"; // 默认黑色
        }

        /// <summary>
        /// 从表格XML中提取合并单元格信息（用于拆分）
        /// </summary>
        public static List<(int row, int col, int rowSpan, int colSpan)> ExtractMergedCellsFromXml(
            XElement tableElement,
            XNamespace w)
        {
            var mergedCells = new List<(int row, int col, int rowSpan, int colSpan)>();

            try
            {
                // 获取所有行
                var xmlRows = tableElement.Elements(w + "tr").ToList();
                
                // 预先计算所有纵向合并的rowspan
                var rowspanMap = CalculateAllRowspansFromXml(xmlRows, w);
                
                // 遍历每一行
                int currentRowIndex = 0; // 0-based，转换为1-based时需要+1
                foreach (var row in xmlRows)
                {
                    int currentLogicalColIndex = 0; // 逻辑列索引（考虑gridSpan）
                    var cells = row.Elements(w + "tc").ToList();
                    
                    foreach (var cell in cells)
                    {
                        var tcPr = cell.Element(w + "tcPr");
                        if (tcPr == null)
                        {
                            currentLogicalColIndex += 1; // 默认colSpan=1
                            continue;
                        }

                        int rowSpan = 1;
                        int colSpan = 1;
                        bool hasMerge = false;

                        // 检查横向合并（gridSpan）
                        var gridSpan = tcPr.Element(w + "gridSpan");
                        if (gridSpan != null)
                        {
                            string gridSpanVal = GetValAttributeFromXml(gridSpan);
                            if (!string.IsNullOrEmpty(gridSpanVal) && int.TryParse(gridSpanVal, out int span))
                            {
                                colSpan = span;
                                hasMerge = true;
                            }
                        }
                        
                        // 检查纵向合并（vMerge）
                        var vMerge = tcPr.Element(w + "vMerge");
                        if (vMerge != null)
                        {
                            string vMergeVal = GetValAttributeFromXml(vMerge);
                            if (vMergeVal == "restart")
                            {
                                // 从rowspanMap获取行跨度
                                if (rowspanMap.TryGetValue((currentRowIndex, currentLogicalColIndex), out int calculatedRowspan))
                                {
                                    rowSpan = calculatedRowspan;
                                    hasMerge = true;
                                }
                            }
                            else
                            {
                                // 这是纵向合并的延续单元格，跳过
                                currentLogicalColIndex += colSpan;
                                continue;
                            }
                        }

                        // 如果单元格被合并（rowSpan > 1 或 colSpan > 1），需要拆分
                        if (hasMerge && (rowSpan > 1 || colSpan > 1))
                        {
                            int row1Based = currentRowIndex + 1;
                            int col1Based = currentLogicalColIndex + 1;
                            mergedCells.Add((row1Based, col1Based, rowSpan, colSpan));
                        }

                        // 更新逻辑列索引
                        currentLogicalColIndex += colSpan;
                    }

                    currentRowIndex++;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 从XML提取合并单元格信息失败：{ex.Message}");
            }

            return mergedCells;
        }
    }
}

