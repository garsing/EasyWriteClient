using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 从XML配置文件创建Word表格的工具。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_CreateTableFromXmlTool
    {
        /// <summary>
        /// XML配置的数据结构
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
        /// 表格样式配置（可选，无 Style 标签时不设置）
        /// </summary>
        public class TableStyleConfig
        {
            /// <summary>表格样式类型，如 ThreeLine（三线表）</summary>
            public string TableStyle { get; set; }
            /// <summary>表头行数，控制三线表第二条线位置，默认1</summary>
            public int HeaderRows { get; set; } = 1;
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
            /// <summary>表格样式配置（三线表等），可选</summary>
            public TableStyleConfig StyleConfig { get; set; }
        }

        /// <summary>
        /// 注册工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_create_table_from_xml"] = async (args) =>
            {
                var phaseSw = Stopwatch.StartNew();
                void LogPhase(string step) =>
                    System.Diagnostics.Debug.WriteLine($"[CreateTable] +{phaseSw.ElapsedMilliseconds}ms {step}");

                try
                {
                    LogPhase("开始");
                    System.Diagnostics.Debug.WriteLine("[DEBUG] 开始创建表格");

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return resolveError;
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[F_create_table_from_xml] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath xmlResolved, out string pathError, "path", "filename"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    string filename = xmlResolved.Display;
                    var xmlRead = await FilePathResolver.ReadAsync(xmlResolved).ConfigureAwait(false);
                    if (!xmlRead.Success)
                    {
                        return new ToolResult { Success = false, Error = xmlRead.Error };
                    }

                    string xmlFilePath = xmlResolved.LocalPath;

                    // 解析XML
                    var parseResult = ParseXmlFile(xmlFilePath);
                    LogPhase($"XML 解析完成 rows={parseResult.Config?.Table?.Rows} cols={parseResult.Config?.Table?.Cols}");
                    if (parseResult.Config == null)
                    {
                        return new ToolResult { Success = false, Error = $"XML解析失败: {parseResult.Error}" };
                    }

                    // 验证配置
                    var validationResult = ValidateConfig(parseResult.Config);
                    if (!validationResult.IsValid)
                    {
                        return new ToolResult { Success = false, Error = $"配置验证失败: {validationResult.Error}" };
                    }

                    if (args.ContainsKey("position") && args["position"] != null
                        && !string.IsNullOrWhiteSpace(args["position"].ToString()))
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "已不再支持顶层 position。正文最前用 target.position=start；句/表/图/图片后用 target.codes / table / chart / image；省略 target 则在当前光标处建表。"
                        };
                    }

                    RefreshDocumentForCreateTable(document, args);

                    var locate = InsertTargetLocatorHelper.TryResolveInsertRange(document, args);
                    if (!locate.Success)
                    {
                        return new ToolResult { Success = false, Error = locate.Error };
                    }

                    var createResult = CreateTableAtLocation(document, parseResult.Config, locate.Range);
                    LogPhase($"CreateTableAtLocation success={createResult.Success}");
                    if (!createResult.Success)
                    {
                        return new ToolResult { Success = false, Error = $"表格创建失败: {createResult.Error}" };
                    }

                    var occupied = new HashSet<string>(DocumentState.GetTableIdOrder() ?? new List<string>());
                    string stampedId = TableTitleStampHelper.NewRandomTableId(occupied);
                    if (!string.IsNullOrEmpty(stampedId))
                    {
                        if (!TableTitleStampHelper.TryWriteStamp(createResult.Table, stampedId))
                        {
                            System.Diagnostics.Debug.WriteLine("[CreateTable] 立刻写 Title 失败，将由 ReadWord O7 兜底");
                        }
                    }

                    var warnings = new List<string>();

                    // 应用格式
                    var formatResult = ApplyTableFormatting(createResult.Table, parseResult.Config.Table);
                    LogPhase($"ApplyTableFormatting success={formatResult.Success}");
                    if (!formatResult.Success)
                    {
                        warnings.Add($"格式应用失败: {formatResult.Error}");
                    }
                    else if (!string.IsNullOrEmpty(formatResult.Error))
                    {
                        warnings.Add($"格式应用警告: {formatResult.Error}");
                    }

                    // 应用表格样式（三线表等）
                    var styleResult = ApplyTableStyle(createResult.Table, parseResult.Config.Table.StyleConfig);
                    if (!styleResult.Success)
                    {
                        warnings.Add($"表格样式应用失败: {styleResult.Error}");
                    }

                    // 填充数据
                    var fillDataResult = FillTableData(createResult.Table, parseResult.Config.Data, parseResult.Config.Table.TableFontColor);
                    LogPhase($"FillTableData success={fillDataResult.Success}");
                    if (!fillDataResult.Success)
                    {
                        warnings.Add($"数据填充失败: {fillDataResult.Error}");
                    }
                    else if (!string.IsNullOrEmpty(fillDataResult.Error))
                    {
                        warnings.Add($"数据填充警告: {fillDataResult.Error}");
                    }

                    // 合并单元格
                    var mergeResult = ApplyCellMerging(createResult.Table, parseResult.Config.Merge);
                    LogPhase($"ApplyCellMerging success={mergeResult.Success}");
                    if (!mergeResult.Success)
                    {
                        warnings.Add($"单元格合并失败: {mergeResult.Error}");
                    }
                    else if (!string.IsNullOrEmpty(mergeResult.Error))
                    {
                        warnings.Add($"单元格合并警告: {mergeResult.Error}");
                    }

                    LogPhase("工具即将返回 ToolResult");

                    string resolvedTableId = null;
                    try
                    {
                        int tableStart = createResult.Table.Range.Start;
                        WordReader.ReadWord(document);

                        List<Word.Table> allTables = TableResolveHelper.GetAllTablesInOrder(document);
                        int tableIndex = allTables.FindIndex(t => t.Range.Start == tableStart);

                        var idOrder = DocumentState.GetTableIdOrder();
                        if (tableIndex >= 0 && tableIndex < idOrder.Count)
                        {
                            resolvedTableId = idOrder[tableIndex];
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 建表 table_id={resolvedTableId}（Range.Start={tableStart}, index={tableIndex}）");
                        }
                        else
                        {
                            warnings.Add("未能解析 table_id（表格索引与编号映射不一致）");
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"未能解析 table_id: {ex.Message}");
                    }

                    string message = InsertTargetLocatorHelper.AppendDocumentStartMessage(
                        "表格创建成功",
                        locate.Spec);
                    if (warnings.Count > 0)
                    {
                        message += $"，但存在以下警告: {string.Join("; ", warnings)}";
                    }

                    LogPhase($"完成 messageLen={message.Length}");

                    Word.Application app = wordApplication as Word.Application;
                    if (app != null && createResult.Table != null)
                    {
                        PostModifyNavigateHelper.NavigateAfterInsertObject(
                            app, createResult.Table.Range, "create_table");
                    }

                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            table_created = true,
                            table_id = resolvedTableId,
                            rows = parseResult.Config.Table.Rows,
                            columns = parseResult.Config.Table.Cols,
                            xml_file = filename,
                            message = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// XML解析结果
        /// </summary>
        private class ParseResult
        {
            public TableConfig Config { get; set; }
            public string Error { get; set; }
        }

        /// <summary>
        /// 解析XML文件
        /// </summary>
        private static ParseResult ParseXmlFile(string filePath)
        {
            try
            {
                XDocument doc = XDocument.Load(filePath);
                XElement root = doc.Root;

                if (root == null || root.Name != "TableConfig")
                {
                    return new ParseResult { Config = null, Error = "XML根元素必须是TableConfig" };
                }

                var generalElement = root.Element("General");
                if (generalElement?.Element("ColWidths") != null)
                {
                    return new ParseResult
                    {
                        Config = null,
                        Error = "TableConfig 布局无效：ColWidths 须位于 Table/Properties 下，请重新提取表格格式",
                    };
                }

                var tableElementForLayout = root.Element("Table");
                if (tableElementForLayout?.Element("ColWidths") != null)
                {
                    return new ParseResult
                    {
                        Config = null,
                        Error = "TableConfig 布局无效：ColWidths 须位于 Table/Properties 下，请重新提取表格格式",
                    };
                }

                if (tableElementForLayout?.Element("Properties")?.Element("Style") != null)
                {
                    return new ParseResult
                    {
                        Config = null,
                        Error = "TableConfig 布局无效：Style 须位于 General 下，请重新提取表格格式",
                    };
                }

                var config = new TableConfig();

                // 解析General（可迁移格式：字体 + Style）
                string tableFontColor = "";
                TableStyleConfig styleConfig = null;

                if (generalElement != null)
                {
                    tableFontColor = generalElement.Element("TableFontColor")?.Value ?? "";

                    var styleElement = generalElement.Element("Style");
                    if (styleElement != null)
                    {
                        string tableStyle = styleElement.Element("TableStyle")?.Value?.Trim() ?? "";
                        int headerRows = int.TryParse(styleElement.Element("HeaderRows")?.Value, out int hr) ? hr : 1;
                        if (!string.IsNullOrEmpty(tableStyle))
                        {
                            styleConfig = new TableStyleConfig
                            {
                                TableStyle = tableStyle,
                                HeaderRows = headerRows,
                            };
                        }
                    }
                }

                List<float> colWidths = new List<float>();

                // 解析Table
                var tableElement = root.Element("Table");
                if (tableElement != null)
                {
                    var propertiesElement = tableElement.Element("Properties");
                    if (propertiesElement != null)
                    {
                        var colWidthsElement = propertiesElement.Element("ColWidths");
                        if (colWidthsElement != null)
                        {
                            foreach (var widthElement in colWidthsElement.Elements("float"))
                            {
                                if (float.TryParse(widthElement.Value, out float width))
                                {
                                    colWidths.Add(width);
                                }
                            }
                        }
                    }

                    config.Table = new TableInfo
                    {
                        ColWidths = colWidths,
                        TableFontColor = tableFontColor,
                        StyleConfig = styleConfig,
                        Style = styleConfig?.TableStyle ?? "",
                    };

                    // 解析 Data（HTML 省略写法；忽略 Properties/Cols/Merge）
                    var dataElement = tableElement.Element("Data");
                    if (dataElement != null)
                    {
                        if (TableMarkupTags.TryGetLegacyGridError(dataElement, out string legacyError))
                        {
                            return new ParseResult { Config = null, Error = legacyError };
                        }

                        var visibleRows = new List<IList<TableHtmlVisibleCell>>();
                        foreach (var rowElement in TableMarkupTags.GridRows(dataElement))
                        {
                            var visible = new List<TableHtmlVisibleCell>();
                            foreach (var cellElement in TableMarkupTags.GridCells(rowElement))
                            {
                                string cellValue;
                                if (cellElement.HasElements)
                                {
                                    cellValue = string.Join("", cellElement.Nodes().Select(n => n.ToString()));
                                }
                                else
                                {
                                    cellValue = cellElement.Value ?? "";
                                }

                                int rowspan = 1;
                                int colspan = 1;
                                string rowspanStr = cellElement.Attribute("rowspan")?.Value;
                                string colspanStr = cellElement.Attribute("colspan")?.Value;
                                if (!string.IsNullOrEmpty(rowspanStr) && int.TryParse(rowspanStr, out int rspan))
                                {
                                    rowspan = rspan;
                                }

                                if (!string.IsNullOrEmpty(colspanStr) && int.TryParse(colspanStr, out int cspan))
                                {
                                    colspan = cspan;
                                }

                                visible.Add(new TableHtmlVisibleCell
                                {
                                    Text = cellValue,
                                    RowSpan = rowspan,
                                    ColSpan = colspan,
                                });
                            }

                            visibleRows.Add(visible);
                        }

                        if (!TableHtmlGrid.TryExpand(visibleRows, out TableHtmlGridResult grid) || !grid.Ok)
                        {
                            return new ParseResult
                            {
                                Config = null,
                                Error = string.IsNullOrEmpty(grid.Error) ? "表格 Data 铺格失败" : grid.Error,
                            };
                        }

                        config.Data = grid.RectangularData;
                        config.Merge = grid.Merge;
                        config.Table.Rows = grid.RowCount;
                        config.Table.Cols = grid.ColCount;

                        if (config.Table.StyleConfig != null)
                        {
                            config.Table.StyleConfig.HeaderRows = Math.Max(
                                1,
                                Math.Min(config.Table.StyleConfig.HeaderRows, config.Table.Rows));
                        }
                    }
                }

                return new ParseResult { Config = config, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"XML解析失败: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细信息: {ex.InnerException.Message}";
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] XML解析异常: {errorMsg}");
                return new ParseResult { Config = null, Error = errorMsg };
            }
        }

        /// <summary>
        /// 验证结果
        /// </summary>
        private class ValidationResult
        {
            public bool IsValid { get; set; }
            public string Error { get; set; }
        }

        /// <summary>
        /// 验证配置
        /// </summary>
        private static ValidationResult ValidateConfig(TableConfig config)
        {
            var errors = new List<string>();

            if (config == null)
            {
                return new ValidationResult { IsValid = false, Error = "配置对象为空" };
            }

            // Location 已移除，不再需要验证

            if (config.Table == null)
            {
                errors.Add("缺少Table配置");
            }
            else
            {
                if (config.Table.Rows <= 0)
                {
                    errors.Add($"Table.Rows值无效: {config.Table.Rows}，必须大于0");
                }

                if (config.Table.Cols <= 0)
                {
                    errors.Add($"Table.Cols值无效: {config.Table.Cols}，必须大于0");
                }

                if (config.Table.ColWidths != null && config.Table.ColWidths.Count > 0)
                {
                    for (int i = 0; i < config.Table.ColWidths.Count; i++)
                    {
                        if (config.Table.ColWidths[i] <= 0)
                        {
                            errors.Add($"Table.ColWidths[{i}]值无效: {config.Table.ColWidths[i]}，必须大于0");
                        }
                    }
                }
            }

            if (config.Data == null)
            {
                errors.Add("缺少Data配置");
            }
            else
            {
                if (config.Table != null)
                {
                    if (config.Data.Count != config.Table.Rows)
                    {
                        errors.Add($"Data行数({config.Data.Count})与Table.Rows({config.Table.Rows})不匹配");
                    }

                    for (int r = 0; r < config.Data.Count; r++)
                    {
                        if (config.Data[r] == null)
                        {
                            errors.Add($"Data第{r + 1}行为空");
                        }
                        else if (config.Table.Cols > 0 && config.Data[r].Count != config.Table.Cols)
                        {
                            errors.Add($"Data第{r + 1}行列数({config.Data[r].Count})与Table.Cols({config.Table.Cols})不匹配");
                        }
                    }
                }
            }

            if (config.Merge != null && config.Merge.Count > 0)
            {
                for (int i = 0; i < config.Merge.Count; i++)
                {
                    if (config.Merge[i] == null)
                    {
                        errors.Add($"Merge第{i + 1}项为空");
                    }
                    else if (config.Merge[i].Count != 4)
                    {
                        errors.Add($"Merge第{i + 1}项元素数量({config.Merge[i].Count})不正确，应为4个元素[startRow, startCol, endRow, endCol]");
                    }
                    else if (config.Table != null)
                    {
                        int startRow = config.Merge[i][0];
                        int startCol = config.Merge[i][1];
                        int endRow = config.Merge[i][2];
                        int endCol = config.Merge[i][3];

                        if (startRow < 1 || startRow > config.Table.Rows)
                        {
                            errors.Add($"Merge第{i + 1}项startRow值({startRow})超出范围[1, {config.Table.Rows}]");
                        }
                        if (startCol < 1 || startCol > config.Table.Cols)
                        {
                            errors.Add($"Merge第{i + 1}项startCol值({startCol})超出范围[1, {config.Table.Cols}]");
                        }
                        if (endRow < 1 || endRow > config.Table.Rows)
                        {
                            errors.Add($"Merge第{i + 1}项endRow值({endRow})超出范围[1, {config.Table.Rows}]");
                        }
                        if (endCol < 1 || endCol > config.Table.Cols)
                        {
                            errors.Add($"Merge第{i + 1}项endCol值({endCol})超出范围[1, {config.Table.Cols}]");
                        }
                        if (startRow > endRow)
                        {
                            errors.Add($"Merge第{i + 1}项startRow({startRow})大于endRow({endRow})");
                        }
                        if (startCol > endCol)
                        {
                            errors.Add($"Merge第{i + 1}项startCol({startCol})大于endCol({endCol})");
                        }
                    }
                }
            }

            if (errors.Count > 0)
            {
                return new ValidationResult { IsValid = false, Error = string.Join("; ", errors) };
            }

            return new ValidationResult { IsValid = true, Error = null };
        }

        /// <summary>
        /// 创建结果
        /// </summary>
        private class CreateResult
        {
            public bool Success { get; set; }
            public Word.Table Table { get; set; }
            public string Error { get; set; }
        }

        /// <summary>
        /// 创建表格
        /// </summary>
        private static void RefreshDocumentForCreateTable(Word.Document document, Dictionary<string, object> args)
        {
            Dictionary<string, object> targetDict = InsertTargetLocatorHelper.TryGetTargetDict(args);
            if (targetDict == null)
            {
                return;
            }

            bool hasCodes = targetDict.ContainsKey("codes") && targetDict["codes"] != null
                && !string.IsNullOrWhiteSpace(targetDict["codes"].ToString());
            try
            {
                if (hasCodes)
                {
                    WordDocumentExtractor.ProcessDocument(
                        document, ProcessDocumentOptions.ForProcessActions("create_table"));
                }
                else
                {
                    WordReader.ReadWord(document);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CreateTable] 定位前刷新失败: " + ex.Message);
            }
        }

        private static CreateResult CreateTableAtLocation(Word.Document document, TableConfig config, Word.Range insertRange)
        {
            try
            {
                if (insertRange == null)
                {
                    return new CreateResult { Success = false, Table = null, Error = "插入位置无效" };
                }

                Word.Table newTable = document.Tables.Add(insertRange, config.Table.Rows, config.Table.Cols);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 表格创建成功: {config.Table.Rows}x{config.Table.Cols}");

                return new CreateResult { Success = true, Table = newTable, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"表格创建异常: {ex.Message}";
                if (ex.InnerException != null)
                {
                    errorMsg += $"\n详细信息: {ex.InnerException.Message}";
                }
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new CreateResult { Success = false, Table = null, Error = errorMsg };
            }
        }

        /// <summary>
        /// 应用表格格式
        /// </summary>
        private static CreateResult ApplyTableFormatting(Word.Table table, TableInfo tableInfo)
        {
            var errors = new List<string>();

            try
            {
                // 设置列宽（使用 TableOptimizer 进行宽度限制）
                if (tableInfo.ColWidths != null && tableInfo.ColWidths.Count > 0)
                {
                    float sumCfg = 0f;
                    for (int k = 0; k < tableInfo.ColWidths.Count; k++)
                    {
                        if (tableInfo.ColWidths[k] > 0)
                            sumCfg += tableInfo.ColWidths[k];
                    }

                    float wMaxPage = TableOptimizer.GetAvailablePageWidthPoints(table);
                    TableOptimizer.LogTune(
                        $"F_CreateTableFromXml.ApplyTableFormatting: 写入 XML 列宽前 配置列宽之和={sumCfg:F1}pt, W_max={wMaxPage:F1}pt, 表格列数={table.Columns.Count}, ColWidths 条目数={tableInfo.ColWidths.Count}（注意：此时尚未填充单元格文本）");

                    // 先应用原始列宽，然后由 TableOptimizer 进行宽度限制
                    int validColCount = Math.Min(tableInfo.ColWidths.Count, table.Columns.Count);
                    for (int i = 0; i < validColCount; i++)
                    {
                        try
                        {
                            if (tableInfo.ColWidths[i] <= 0)
                            {
                                errors.Add($"列{i + 1}宽度值无效: {tableInfo.ColWidths[i]}，已跳过");
                                continue;
                            }
                            
                            table.Columns[i + 1].Width = tableInfo.ColWidths[i];
                        }
                        catch (Exception ex)
                        {
                            string errorMsg = $"设置列{i + 1}宽度失败: {ex.Message}";
                            errors.Add(errorMsg);
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                        }
                    }

                    // 使用页面实测可用宽度（与 TableOptimizer 设计一致，不再写死 400）
                    TableOptimizer.LimitTableWidth(table, tableInfo.ColWidths, null);
                    TableOptimizer.LogTune(
                        $"F_CreateTableFromXml.ApplyTableFormatting: LimitTableWidth 之后 Word 列宽之和={TableOptimizer.GetTableColumnsWidthSum(table):F1}pt");
                }
                else
                {
                    TableOptimizer.LogTune(
                        "F_CreateTableFromXml.ApplyTableFormatting: 未提供 ColWidths，跳过列宽写入与 LimitTableWidth");
                }

                // 应用样式
                if (!string.IsNullOrEmpty(tableInfo.Style))
                {
                    try
                    {
                        object styleObj = tableInfo.Style;
                        table.set_Style(ref styleObj);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用样式: {tableInfo.Style}");
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"样式'{tableInfo.Style}'应用失败: {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                        // 尝试常见样式
                        bool fallbackSuccess = false;
                        string[] commonStyles = { "Table Grid", "Table Normal" };
                        foreach (string styleName in commonStyles)
                        {
                            try
                            {
                                object styleObj = styleName;
                                table.set_Style(ref styleObj);
                                System.Diagnostics.Debug.WriteLine($"[DEBUG] 使用备用样式: {styleName}");
                                fallbackSuccess = true;
                                break;
                            }
                            catch { }
                        }
                        if (!fallbackSuccess)
                        {
                            errors.Add("无法应用任何备用样式");
                        }
                    }
                }

                // 设置边框
                try
                {
                    table.Borders.Enable = 1;
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置边框失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                // 与 TableOptimizer.DefaultOptimizationMaxFontSize / OptimizationConfig.MaxFontSize 默认一致，避免填数阶段过小再被优化骤然拉大
                try
                {
                    table.Range.Font.Size = TableOptimizer.DefaultOptimizationMaxFontSize;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 设置表格默认字体大小: {TableOptimizer.DefaultOptimizationMaxFontSize}pt（与 OptimizeTable MaxFontSize 默认一致）");
                }
                catch (Exception ex)
                {
                    string errorMsg = $"设置默认字体大小失败: {ex.Message}";
                    errors.Add(errorMsg);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                }

                if (errors.Count > 0)
                {
                    return new CreateResult { Success = true, Table = table, Error = string.Join("; ", errors) };
                }

                return new CreateResult { Success = true, Table = table, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"格式应用异常: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new CreateResult { Success = false, Table = table, Error = errorMsg };
            }
        }

        /// <summary>
        /// 应用表格样式（三线表等）
        /// </summary>
        private static CreateResult ApplyTableStyle(Word.Table table, TableStyleConfig styleConfig)
        {
            if (styleConfig == null || string.IsNullOrEmpty(styleConfig.TableStyle))
            {
                return new CreateResult { Success = true, Table = table, Error = null };
            }

            var styleResult = TableStyleApplyHelper.Apply(
                table, styleConfig.TableStyle, styleConfig.HeaderRows);
            return new CreateResult
            {
                Success = styleResult.Success,
                Table = table,
                Error = styleResult.Error,
            };
        }

        /// <summary>
        /// 填充表格数据
        /// </summary>
        private static CreateResult FillTableData(Word.Table table, List<List<string>> data, string defaultFontColor)
        {
            var errors = new List<string>();

            for (int r = 0; r < data.Count && r < table.Rows.Count; r++)
            {
                for (int c = 0; c < data[r].Count && c < table.Columns.Count; c++)
                {
                    try
                    {
                        Word.Cell cell = table.Cell(r + 1, c + 1);
                        string cellData = data[r][c] ?? "";

                        System.Diagnostics.Debug.WriteLine($"[FillTableData] 📍 处理单元格({r + 1},{c + 1}): '{cellData}'");

                        // 优先级：HTML样式 > Table默认样式
                        if (cellData.Contains("<"))
                        {
                            System.Diagnostics.Debug.WriteLine($"[FillTableData] 🔍 检测到HTML内容，准备调用InjectHtml");
                            // HTML内容：让MSHTML引擎处理样式，HTML样式优先
                            var htmlResult = TableCellContentWriter.InjectHtml(cell.Range, cellData);
                            if (!htmlResult.Success)
                            {
                                errors.Add($"单元格({r + 1},{c + 1})HTML注入失败: {htmlResult.Error}");
                            }
                            else if (!string.IsNullOrEmpty(htmlResult.Warning))
                            {
                                errors.Add($"单元格({r + 1},{c + 1})HTML注入警告: {htmlResult.Warning}");
                            }
                            System.Diagnostics.Debug.WriteLine($"[FillTableData] ✅ HTML内容处理完成: {cellData}, 结果={htmlResult.Success}");
                        }
                        else
                        {
                            // 纯文本内容：设置文本，然后应用Table默认字体颜色
                            cell.Range.Text = cellData;

                            if (!string.IsNullOrEmpty(defaultFontColor))
                            {
                                try
                                {
                                    cell.Range.Font.Color = (Word.WdColor)Convert.ToInt32(defaultFontColor, 16);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 应用默认字体颜色: {defaultFontColor}");
                                }
                                catch (Exception ex)
                                {
                                    string errorMsg = $"单元格({r + 1},{c + 1})设置默认字体颜色失败: {ex.Message}（颜色值格式可能不正确，应为16进制字符串，如'000000'）";
                                    errors.Add(errorMsg);
                                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // 可能是被合并的单元格，记录警告
                        string errorMsg = $"单元格({r + 1},{c + 1})填充失败: {ex.Message}（可能是已合并的单元格）";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }
            }

            if (errors.Count > 0)
            {
                return new CreateResult { Success = true, Table = table, Error = string.Join("; ", errors) };
            }

            return new CreateResult { Success = true, Table = table, Error = null };
        }

        /// <summary>
        /// 列合并操作信息
        /// </summary>
        private class ColumnMergeInfo
        {
            public int StartRow { get; set; }
            public int EndRow { get; set; }
            public int StartCol { get; set; } // 原始列索引
            public int EndCol { get; set; }   // 原始列索引
            public int MergedColCount { get; set; } // 被合并的列数（endCol - startCol）
        }

        /// <summary>
        /// 合并单元格
        /// 维护索引映射表，处理列合并导致的列索引变化
        /// </summary>
        private static CreateResult ApplyCellMerging(Word.Table table, List<List<int>> mergeConfig)
        {
            if (mergeConfig == null || mergeConfig.Count == 0)
                return new CreateResult { Success = true, Table = table, Error = null };

            var errors = new List<string>();

            try
            {
                // 维护每行的列合并操作列表
                var columnMergeOps = new Dictionary<int, List<ColumnMergeInfo>>(); // 行号 -> 列合并操作列表

                for (int i = 0; i < mergeConfig.Count; i++)
                {
                    var merge = mergeConfig[i];
                    if (merge == null)
                    {
                        errors.Add($"Merge第{i + 1}项为空");
                        continue;
                    }

                    if (merge.Count != 4)
                    {
                        errors.Add($"Merge第{i + 1}项元素数量({merge.Count})不正确，应为4个元素[startRow, startCol, endRow, endCol]");
                        continue;
                    }

                    int originalStartRow = merge[0], originalStartCol = merge[1];
                    int originalEndRow = merge[2], originalEndCol = merge[3];

                    try
                    {
                        // 验证范围
                        if (originalStartRow < 1 || originalStartRow > table.Rows.Count)
                        {
                            errors.Add($"Merge第{i + 1}项startRow({originalStartRow})超出表格行数范围[1, {table.Rows.Count}]");
                            continue;
                        }
                        if (originalEndRow < 1 || originalEndRow > table.Rows.Count)
                        {
                            errors.Add($"Merge第{i + 1}项endRow({originalEndRow})超出表格行数范围[1, {table.Rows.Count}]");
                            continue;
                        }

                        // 判断合并类型
                        bool isRowMerge = originalStartRow != originalEndRow; // 行合并
                        bool isColMerge = originalStartCol != originalEndCol; // 列合并

                        // 通过列合并操作列表转换列索引
                        int actualStartCol = GetActualColumnIndex(originalStartRow, originalStartCol, columnMergeOps);
                        int actualEndCol = GetActualColumnIndex(originalEndRow, originalEndCol, columnMergeOps);

                        // 验证转换后的列索引
                        if (actualStartCol < 1 || actualStartCol > table.Columns.Count)
                        {
                            errors.Add($"Merge第{i + 1}项转换后的startCol({actualStartCol})超出表格列数范围[1, {table.Columns.Count}]");
                            continue;
                        }
                        if (actualEndCol < 1 || actualEndCol > table.Columns.Count)
                        {
                            errors.Add($"Merge第{i + 1}项转换后的endCol({actualEndCol})超出表格列数范围[1, {table.Columns.Count}]");
                            continue;
                        }

                        System.Diagnostics.Debug.WriteLine($"[DEBUG] Merge第{i + 1}项: 原始({originalStartRow},{originalStartCol})->({originalEndRow},{originalEndCol}), 转换后({originalStartRow},{actualStartCol})->({originalEndRow},{actualEndCol}), 行合并={isRowMerge}, 列合并={isColMerge}");

                        // 执行合并
                        Word.Cell startCell = table.Cell(originalStartRow, actualStartCol);
                        Word.Cell endCell = table.Cell(originalEndRow, actualEndCol);
                        startCell.Merge(endCell);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 合并单元格成功: ({originalStartRow},{actualStartCol})->({originalEndRow},{actualEndCol})");

                        // 如果是列合并，记录列合并操作
                        if (isColMerge)
                        {
                            RecordColumnMerge(originalStartRow, originalEndRow, originalStartCol, originalEndCol, columnMergeOps);
                        }
                    }
                    catch (Exception ex)
                    {
                        string errorMsg = $"Merge第{i + 1}项合并失败({originalStartRow},{originalStartCol})->({originalEndRow},{originalEndCol}): {ex.Message}";
                        errors.Add(errorMsg);
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    }
                }

                if (errors.Count > 0)
                {
                    return new CreateResult { Success = true, Table = table, Error = string.Join("; ", errors) };
                }

                return new CreateResult { Success = true, Table = table, Error = null };
            }
            catch (Exception ex)
            {
                string errorMsg = $"合并异常: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                return new CreateResult { Success = false, Table = table, Error = errorMsg };
            }
        }

        /// <summary>
        /// 获取实际的列索引（通过列合并操作列表计算）
        /// </summary>
        private static int GetActualColumnIndex(int row, int originalCol, Dictionary<int, List<ColumnMergeInfo>> columnMergeOps)
        {
            // 如果该行没有列合并操作，直接返回原始列索引
            if (!columnMergeOps.ContainsKey(row) || columnMergeOps[row].Count == 0)
            {
                return originalCol;
            }

            var mergeOps = columnMergeOps[row];
            int actualCol = originalCol;
            int totalOffset = 0;

            // 遍历该行的所有列合并操作，计算累积偏移量
            // 按照原始列索引排序，确保按顺序处理
            var sortedOps = mergeOps.OrderBy(op => op.StartCol).ToList();

            foreach (var mergeOp in sortedOps)
            {
                // 如果原始列索引在合并范围内，说明这个列被合并了
                if (originalCol >= mergeOp.StartCol && originalCol <= mergeOp.EndCol)
                {
                    // 被合并的列，实际列索引 = 合并起始列的实际列索引
                    // 需要递归计算合并起始列的实际列索引（考虑之前的合并操作）
                    actualCol = GetActualColumnIndex(row, mergeOp.StartCol, columnMergeOps);
                    return actualCol;
                }
                // 如果原始列索引在合并范围之后，需要减去被合并的列数
                else if (originalCol > mergeOp.EndCol)
                {
                    // 累积偏移量：减去被合并的列数
                    totalOffset += mergeOp.MergedColCount;
                }
            }

            // 实际列索引 = 原始列索引 - 累积偏移量
            actualCol = originalCol - totalOffset;
            return actualCol;
        }

        /// <summary>
        /// 记录列合并操作
        /// </summary>
        private static void RecordColumnMerge(int startRow, int endRow, int startCol, int endCol, Dictionary<int, List<ColumnMergeInfo>> columnMergeOps)
        {
            // 列合并会影响从startRow到endRow的所有行
            for (int row = startRow; row <= endRow; row++)
            {
                // 获取或创建该行的列合并操作列表
                if (!columnMergeOps.ContainsKey(row))
                {
                    columnMergeOps[row] = new List<ColumnMergeInfo>();
                }

                var mergeOps = columnMergeOps[row];

                // 创建列合并操作信息
                var mergeInfo = new ColumnMergeInfo
                {
                    StartRow = startRow,
                    EndRow = endRow,
                    StartCol = startCol,
                    EndCol = endCol,
                    MergedColCount = endCol - startCol
                };

                mergeOps.Add(mergeInfo);
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] 记录列合并操作: 行{startRow}-{endRow}, 列{startCol}-{endCol}, 合并列数={endCol - startCol}");
        }

        /// <summary>
        /// 使用自定义HTML解析器注入HTML内容（不使用MSHTML引擎）
        /// </summary>
        private static CreateResult InjectHtml(Word.Range range, string htmlContent)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] InjectHtml: 开始注入HTML内容（使用自定义解析器）");
                System.Diagnostics.Debug.WriteLine($"[DEBUG]   HTML内容：{htmlContent}");
                
                // 解析HTML内容
                var segments = HtmlParser.ParseHtml(htmlContent);
                System.Diagnostics.Debug.WriteLine($"[InjectHtml]   解析到 {segments.Count} 个文本段");
                
                // 检查是否有颜色设置
                int colorCount = segments.Count(s => !string.IsNullOrEmpty(s.Color));
                if (colorCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[InjectHtml]   检测到 {colorCount} 个颜色设置：");
                    foreach (var segment in segments.Where(s => !string.IsNullOrEmpty(s.Color)))
                    {
                        System.Diagnostics.Debug.WriteLine($"[InjectHtml]     文本：'{segment.Text}'，颜色：{segment.Color}");
                    }
                }
                
                // 检查是否有加粗设置
                int boldCount = segments.Count(s => s.IsBold);
                if (boldCount > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[InjectHtml]   ⚠️ 检测到 {boldCount} 个加粗设置：");
                    foreach (var segment in segments.Where(s => s.IsBold))
                    {
                        System.Diagnostics.Debug.WriteLine($"[InjectHtml]     文本：'{segment.Text}'，IsBold={segment.IsBold}");
                    }
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[InjectHtml]   ⚠️ 未检测到任何加粗设置");
                }
                
                // 插入文本段并设置格式
                bool success = HtmlParser.InsertHtmlSegments(range, segments);
                
                if (success)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] ✅ HTML注入成功，最终段落数：{range.Paragraphs.Count}");
                    return new CreateResult { Success = true, Table = null, Error = null };
                }
                else
                {
                    // 降级为纯文本
                    string plainText = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<[^>]+>", "");
                    range.Text = plainText;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] ⚠️ HTML解析失败，已降级为纯文本");
                    return new CreateResult { Success = true, Table = null, Error = "HTML解析失败，已降级为纯文本" };
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] ❌ HTML注入失败：{ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 异常详情：{ex}");
                try
                {
                    string plainText = System.Text.RegularExpressions.Regex.Replace(htmlContent, "<[^>]+>", "");
                    range.Text = plainText;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 已降级为纯文本");
                    return new CreateResult { Success = true, Table = null, Error = $"HTML注入失败，已降级为纯文本: {ex.Message}" };
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] ❌ 降级为纯文本也失败：{ex2.Message}");
                    return new CreateResult { Success = false, Table = null, Error = $"{ex.Message}，降级为纯文本也失败: {ex2.Message}" };
                }
            }
        }

    }
}