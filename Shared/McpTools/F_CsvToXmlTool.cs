using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// XML文件类型枚举
    /// </summary>
    public enum XmlFileType
    {
        TableFormat,  // 表格格式：TableConfig/Table/Data
        ChartFormat   // 图表格式：ChartConfig/Chart/DataTable
    }

    /// <summary>
    /// CSV转XML工具
    /// 将CSV文件转换为XML格式的data节点
    /// </summary>
    public static class F_CsvToXmlTool
    {
        /// <summary>
        /// 注册CSV转XML工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_csv_to_xml"] = async (args) =>
            {
                await Task.CompletedTask; // 确保异步执行

                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] csv_to_xml工具开始执行");

                    string csvFilename = args.ContainsKey("csv_filename") ? args["csv_filename"]?.ToString() : "";
                    string xmlFilename = args.ContainsKey("xml_filename") ? args["xml_filename"]?.ToString() : "";

                    if (string.IsNullOrEmpty(csvFilename))
                    {
                        return new ToolResult { Success = false, Error = "必须提供csv_filename参数（CSV文件名）" };
                    }

                    if (string.IsNullOrEmpty(xmlFilename))
                    {
                        return new ToolResult { Success = false, Error = "必须提供xml_filename参数（XML文件名）" };
                    }

                    if (!UserService.Instance.CheckLoginStatus()
                        || string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
                    {
                        return new ToolResult { Success = false, Error = "用户未登录或工作区未初始化" };
                    }

                    var csvEnsure = await McpToolsHelpers.EnsureWorkspaceFileAsync(csvFilename).ConfigureAwait(false);
                    if (!csvEnsure.success)
                    {
                        return new ToolResult { Success = false, Error = csvEnsure.error };
                    }

                    string csvFilePath = csvEnsure.localPath;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 读取CSV文件: {csvFilePath}");

                    var xmlEnsure = await McpToolsHelpers.EnsureWorkspaceFileAsync(xmlFilename).ConfigureAwait(false);
                    string xmlFilePath = xmlEnsure.success
                        ? xmlEnsure.localPath
                        : WorkspacePathResolver.ResolveWritePath(xmlFilename);

                    Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath));
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 操作XML文件: {xmlFilePath}");

                    // 读取并解析CSV文件
                    var csvData = ParseCsvFile(csvFilePath);
                    if (csvData == null || csvData.Count == 0)
                    {
                        return new ToolResult { Success = false, Error = "CSV文件解析失败或文件为空" };
                    }

                    // 检查XML文件是否存在，不存在则创建新文件
                    if (!xmlEnsure.success && !File.Exists(xmlFilePath))
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 将创建新的XML文件: {xmlFilename}");
                    }

                    // 读取或创建XML文件
                    XDocument xmlDoc;
                    bool xmlFileExists = File.Exists(xmlFilePath);
                    XmlFileType xmlFileType = XmlFileType.TableFormat;

                    if (xmlFileExists)
                    {
                        // 读取现有XML文件并检测类型
                        xmlDoc = XDocument.Load(xmlFilePath);
                        xmlFileType = DetectXmlFileType(xmlDoc);
                    }
                    else
                    {
                        // 创建新的XML文档（默认为表格格式）
                        xmlDoc = new XDocument();
                        xmlFileType = XmlFileType.TableFormat;
                    }

                    // 根据XML文件类型更新或添加Data节点
                    if (xmlFileType == XmlFileType.TableFormat)
                    {
                        // 表格格式：TableConfig/Table/Data
                        ImportDataToTableFormat(xmlDoc, csvData);
                    }
                    else if (xmlFileType == XmlFileType.ChartFormat)
                    {
                        // 图表格式：ChartConfig/Chart/Data
                        ImportDataToChartFormat(xmlDoc, csvData);
                    }
                    else
                    {
                        return new ToolResult { Success = false, Error = "无法识别的XML文件格式" };
                    }

                    // 保存XML文件
                    xmlDoc.Save(xmlFilePath);

                    // 上传生成的XML文件到用户工作目录
                    bool uploadSuccess = await McpToolsHelpers.UploadWorkspaceFileAsync(xmlFilePath, xmlFilename);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] XML文件上传结果: {uploadSuccess}");

                    // 获取文件信息
                    FileInfo xmlFileInfo = new FileInfo(xmlFilePath);

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            csv_filename = csvFilename,
                            xml_filename = xmlFilename,
                            csv_file_path = csvFilePath,
                            xml_file_path = xmlFilePath,
                            csv_rows = csvData.Count,
                            csv_columns = csvData.Count > 0 ? csvData[0].Count : 0,
                            xml_file_created = !xmlFileExists,
                            xml_file_type = xmlFileType.ToString(),
                            xml_file_size = xmlFileInfo.Length,
                            modified = xmlFileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            upload_success = uploadSuccess,
                            message = xmlFileExists ?
                                $"成功将CSV文件转换为XML的data节点，已更新 {xmlFilename}（格式：{xmlFileType.ToString()}）{(uploadSuccess ? " 并上传到用户工作目录" : "（上传失败）")}" :
                                $"成功将CSV文件转换为XML的data节点，已创建 {xmlFilename}（格式：{xmlFileType.ToString()}）{(uploadSuccess ? " 并上传到用户工作目录" : "（上传失败）")}"
                        }
                    };

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] csv_to_xml工具执行成功，已处理 {csvData.Count} 行数据");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] csv_to_xml工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"CSV转XML失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// 解析CSV文件
        /// </summary>
        private static List<List<string>> ParseCsvFile(string filePath)
        {
            try
            {
                var lines = File.ReadAllLines(filePath, Encoding.UTF8);
                var result = new List<List<string>>();

                foreach (var line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var row = ParseCsvLine(line);
                    if (row.Count > 0)
                    {
                        result.Add(row);
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] CSV文件解析失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 解析CSV行，支持带引号的字段
        /// </summary>
        private static List<string> ParseCsvLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            var inQuotes = false;
            var quoteChar = '"';

            for (int i = 0; i < line.Length; i++)
            {
                var c = line[i];

                if (!inQuotes)
                {
                    if (c == '"' || c == '\'')
                    {
                        // 开始引号
                        inQuotes = true;
                        quoteChar = c;
                    }
                    else if (c == ',')
                    {
                        // 字段结束
                        result.Add(current.ToString().Trim());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else
                {
                    if (c == quoteChar)
                    {
                        // 检查是否是转义的引号
                        if (i + 1 < line.Length && line[i + 1] == quoteChar)
                        {
                            current.Append(quoteChar);
                            i++; // 跳过下一个引号
                        }
                        else
                        {
                            // 引号结束
                            inQuotes = false;
                        }
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
            }

            // 添加最后一个字段
            result.Add(current.ToString().Trim());

            return result;
        }

        /// <summary>
        /// 检测XML文件的类型
        /// </summary>
        private static XmlFileType DetectXmlFileType(XDocument xmlDoc)
        {
            try
            {
                if (xmlDoc.Root == null)
                {
                    return XmlFileType.TableFormat;
                }

                // 检查根元素名称
                if (xmlDoc.Root.Name == "ChartConfig")
                {
                    return XmlFileType.ChartFormat;
                }
                else if (xmlDoc.Root.Name == "TableConfig")
                {
                    return XmlFileType.TableFormat;
                }

                // 默认认为是表格格式
                return XmlFileType.TableFormat;
            }
            catch
            {
                // 出错时默认为表格格式
                return XmlFileType.TableFormat;
            }
        }

        /// <summary>
        /// 将CSV数据导入到表格格式的XML文件中
        /// </summary>
        private static void ImportDataToTableFormat(XDocument xmlDoc, List<List<string>> csvData)
        {
            try
            {
                XElement root;
                if (xmlDoc.Root == null)
                {
                    root = new XElement("TableConfig");
                    xmlDoc.Add(root);
                }
                else
                {
                    root = xmlDoc.Root;
                }

                // 确保Table节点存在
                XElement tableElement = root.Element("Table");
                if (tableElement == null)
                {
                    tableElement = new XElement("Table");
                    root.Add(tableElement);
                }

                // 创建或更新Data节点
                XElement dataElement = tableElement.Element("Data");
                if (dataElement != null)
                {
                    dataElement.Remove();
                }
                dataElement = new XElement("Data");

                foreach (var row in csvData)
                {
                    var rowElement = new XElement("Row");
                    foreach (var cell in row)
                    {
                        rowElement.Add(new XElement("Cell", cell));
                    }
                    dataElement.Add(rowElement);
                }

                tableElement.Add(dataElement);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 表格格式数据导入失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 将CSV数据导入到图表格式的XML文件中（新格式：DataTable/Header/Rows）
        /// CSV格式：第一列是X值，后续列是各个Series的Y值
        /// 如果CSV第一行看起来像表头（非数字），则将其作为Column名称；否则自动生成Series名称
        /// </summary>
        private static void ImportDataToChartFormat(XDocument xmlDoc, List<List<string>> csvData)
        {
            try
            {
                XElement root;
                if (xmlDoc.Root == null)
                {
                    root = new XElement("ChartConfig");
                    xmlDoc.Add(root);
                }
                else
                {
                    root = xmlDoc.Root;
                }

                // 确保Chart节点存在
                XElement chartElement = root.Element("Chart");
                if (chartElement == null)
                {
                    chartElement = new XElement("Chart");
                    root.Add(chartElement);
                }

                // 检查CSV数据是否有效
                if (csvData == null || csvData.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] CSV数据为空，跳过更新");
                    return;
                }

                // 确定Series数量（CSV列数 - 1，第一列是X值）
                int seriesCount = csvData[0].Count - 1;
                if (seriesCount <= 0)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] CSV数据列数不足，至少需要2列（X列+至少1个Y列）");
                    return;
                }

                // 检查第一行是否是表头（如果第一行的值看起来不像数字，则认为是表头）
                bool hasHeader = false;
                List<string> headerRow = null;
                List<List<string>> dataRows = csvData;

                if (csvData.Count > 0)
                {
                    var firstRow = csvData[0];
                    // 检查第一行除了第一列外，是否有非数字的值（可能是表头）
                    bool firstRowLooksLikeHeader = false;
                    for (int i = 1; i < firstRow.Count; i++)
                    {
                        string value = firstRow[i]?.Trim() ?? "";
                        if (!string.IsNullOrEmpty(value) && !double.TryParse(value, out _))
                        {
                            firstRowLooksLikeHeader = true;
                            break;
                        }
                    }

                    // 如果第一行看起来像表头，或者第一行第一列不是数字，则认为是表头
                    if (firstRowLooksLikeHeader || (firstRow.Count > 0 && !double.TryParse(firstRow[0]?.Trim() ?? "", out _)))
                    {
                        hasHeader = true;
                        headerRow = csvData[0];
                        dataRows = csvData.Skip(1).ToList();
                    }
                }

                // 获取或创建DataTable节点
                XElement dataTableElement = chartElement.Element("DataTable");
                if (dataTableElement == null)
                {
                    dataTableElement = new XElement("DataTable");
                    chartElement.Add(dataTableElement);
                }

                // 创建或更新Header节点
                XElement headerElement = dataTableElement.Element("Header");
                if (headerElement != null)
                {
                    headerElement.Remove();
                }
                headerElement = new XElement("Header");

                // 创建category列（X轴）
                string categoryColumnName = hasHeader && headerRow.Count > 0 ? headerRow[0] : "类别";
                headerElement.Add(new XElement("Column",
                    new XAttribute("type", "category"),
                    categoryColumnName
                ));

                // 获取现有的Column定义（如果存在），以保留属性（如chartType、color、axisY等）
                var existingDataTable = chartElement.Element("DataTable");
                var existingHeader = existingDataTable?.Element("Header");
                var existingColumns = existingHeader?.Elements("Column").Skip(1).ToList() ?? new List<XElement>();

                // 创建value列（Y轴数据系列）
                for (int i = 0; i < seriesCount; i++)
                {
                    string seriesName;
                    if (hasHeader && headerRow.Count > i + 1)
                    {
                        seriesName = headerRow[i + 1] ?? $"系列{i + 1}";
                    }
                    else
                    {
                        seriesName = $"系列{i + 1}";
                    }

                    // 尝试从现有XML中获取该Series的属性，如果不存在则使用默认值
                    XElement existingColumn = existingColumns.Count > i ? existingColumns[i] : null;
                    
                    var columnElement = new XElement("Column",
                        new XAttribute("type", "value")
                    );

                    // 如果现有Column有属性，保留它们；否则使用默认值
                    if (existingColumn != null)
                    {
                        // 保留现有属性
                        foreach (var attr in existingColumn.Attributes())
                        {
                            if (attr.Name != "type") // type属性必须为value
                            {
                                columnElement.SetAttributeValue(attr.Name, attr.Value);
                            }
                        }
                    }
                    else
                    {
                        // 使用默认属性
                        columnElement.SetAttributeValue("axisY", "primary");
                        columnElement.SetAttributeValue("showDataLabels", "false");
                    }

                    // 设置Column的文本内容（系列名称）
                    columnElement.Value = seriesName;
                    headerElement.Add(columnElement);
                }

                dataTableElement.Add(headerElement);

                // 创建或更新Rows节点
                XElement rowsElement = dataTableElement.Element("Rows");
                if (rowsElement != null)
                {
                    rowsElement.Remove();
                }
                rowsElement = new XElement("Rows");

                // 将CSV数据行转换为Row元素（逗号分隔的值）
                foreach (var row in dataRows)
                {
                    if (row.Count > 0)
                    {
                        // 构建逗号分隔的字符串：第一列是X值，后续列是Y值
                        var rowValues = new List<string>();
                        rowValues.Add(row[0] ?? ""); // X值

                        // 添加各个Series的Y值
                        for (int i = 1; i < row.Count && i <= seriesCount; i++)
                        {
                            rowValues.Add(row[i] ?? "");
                        }

                        // 确保所有Series都有值（如果CSV行数据不足，用空字符串填充）
                        while (rowValues.Count <= seriesCount)
                        {
                            rowValues.Add("");
                        }

                        string rowData = string.Join(",", rowValues);
                        rowsElement.Add(new XElement("Row", rowData));
                    }
                }

                dataTableElement.Add(rowsElement);

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 成功更新图表格式数据：{seriesCount}个Series，{dataRows.Count}个数据行");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 图表格式数据导入失败: {ex.Message}");
                throw;
            }
        }

    }
}
