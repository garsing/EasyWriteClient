using System.Threading.Tasks;

namespace WordAddIn1
{
    public sealed class ChartParseResult
    {
        public ChartConfig Config { get; set; }
        public string DataSource { get; set; }
        public string CsvFileName { get; set; }
        public string Error { get; set; }

        public bool Success => Config != null && string.IsNullOrEmpty(Error);
    }

    /// <summary>
    /// ChartConfig XML 解析（与 F_create_chart_from_xml 共用实现）。
    /// </summary>
    public static class ChartConfigXmlParser
    {
        public static async Task<ChartParseResult> ParseAsync(string xmlContent, string filenameOptional = null)
        {
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                return new ChartParseResult { Error = "XML内容为空" };
            }

            var parseResult = await ChartXmlParser.ParseXmlString(xmlContent);
            if (parseResult.Config == null)
            {
                return new ChartParseResult { Error = parseResult.Error ?? "XML解析失败" };
            }

            return new ChartParseResult
            {
                Config = parseResult.Config,
                DataSource = parseResult.DataSource,
                CsvFileName = parseResult.CsvFileName,
                Error = parseResult.Error
            };
        }

        public static async Task<ChartParseResult> ParseFromFileAsync(string filePath)
        {
            try
            {
                string xmlContent = System.IO.File.ReadAllText(filePath, System.Text.Encoding.UTF8);
                return await ParseAsync(xmlContent, filePath);
            }
            catch (System.Exception ex)
            {
                return new ChartParseResult { Error = $"读取XML文件失败: {ex.Message}" };
            }
        }
    }
}
