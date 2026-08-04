using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 测试标题识别工具
    /// 读取Word文档，识别其中的标题，并总结识别结果
    /// </summary>
    public static class F_TestHeadingRecognizer
    {
        /// <summary>
        /// 注册测试标题识别工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_test_heading_recognizer"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    dynamic wordApp = wordApplication;

                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    Word.Document doc = wordApp.ActiveDocument;

                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("=== F_test_heading_recognizer 工具执行 ===");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine($"文档名称: {doc.Name ?? "未命名文档"}");
                    System.Diagnostics.Debug.WriteLine("");

                    var totalStartTime = System.Diagnostics.Stopwatch.StartNew();

                    // 第一步：调用WordDocumentExtractor读取文档
                    System.Diagnostics.Debug.WriteLine("[步骤1] 调用WordDocumentExtractor读取文档...");
                    var step1StartTime = System.Diagnostics.Stopwatch.StartNew();
                    var processingResult = WordDocumentExtractor.ProcessDocument(doc);
                    step1StartTime.Stop();
                    
                    if (processingResult == null)
                    {
                        return new ToolResult { Success = false, Error = "文档处理失败" };
                    }

                    System.Diagnostics.Debug.WriteLine($"[步骤1] ✅ 文档处理完成，处理方式: {processingResult.ProcessingMethod}，耗时: {step1StartTime.ElapsedMilliseconds}ms ({step1StartTime.ElapsedMilliseconds / 1000.0:F2}秒)");

                    // 第二步：按行分类（格式映射已由 ProcessDocument 写入 DocumentState，此处仅复现分类结果用于展示）
                    System.Diagnostics.Debug.WriteLine("[步骤2] 使用 ReadText 进行按行分类摘要...");
                    var step2StartTime = System.Diagnostics.Stopwatch.StartNew();
                    
                    string fullText = processingResult.ReadText;
                    if (string.IsNullOrEmpty(fullText))
                    {
                        return new ToolResult { Success = false, Error = "ReadText 为空，文档可能未正确处理" };
                    }

                    DocumentLineClassificationResult classification = new HeadingRecognizer().ClassifyLinesForFormatExtraction(fullText);
                    List<Dictionary<string, object>> subtypeFormatMapping = DocumentState.SubtypeFormatMapping;
                    
                    step2StartTime.Stop();
                    
                    int headingLineCount = classification.HeadingsByLanguage.Values.Sum(l => l.Count);
                    System.Diagnostics.Debug.WriteLine($"[步骤2] ✅ 标题行 {headingLineCount}，正文候选行 {classification.BodyLines.Count}，耗时: {step2StartTime.ElapsedMilliseconds}ms");

                    // 第三步：输出按行分类摘要
                    System.Diagnostics.Debug.WriteLine("[步骤3] 输出按行分类摘要...");
                    var step3StartTime = System.Diagnostics.Stopwatch.StartNew();
                    
                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("按行分类摘要（标题按语言分组）");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("按行分类摘要（标题按语言分组）");
                    Console.WriteLine("==========================================");
                    
                    foreach (var langKv in classification.HeadingsByLanguage)
                    {
                        string langHeader = $"语言: {langKv.Key}，标题数: {langKv.Value.Count}";
                        System.Diagnostics.Debug.WriteLine(langHeader);
                        Console.WriteLine(langHeader);
                        foreach (Dictionary<string, object> h in langKv.Value)
                        {
                            string lineNum = h.ContainsKey("line") ? h["line"]?.ToString() ?? "" : "";
                            string dst = h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                            string text = h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                            string lineOut = $"  line {lineNum} [{dst}] {text}";
                            System.Diagnostics.Debug.WriteLine(lineOut);
                            Console.WriteLine(lineOut);
                        }
                    }
                    
                    string bodySummary = $"正文候选行数: {classification.BodyLines.Count}";
                    System.Diagnostics.Debug.WriteLine(bodySummary);
                    Console.WriteLine(bodySummary);
                    
                    step3StartTime.Stop();
                    System.Diagnostics.Debug.WriteLine($"[步骤3] ✅ 耗时: {step3StartTime.ElapsedMilliseconds}ms");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("");
                    
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");
                    
                    // 输出JSON格式的映射关系
                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("detailed_subtype 格式映射关系（JSON）");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    
                    string jsonOutput = JsonConvert.SerializeObject(subtypeFormatMapping, Formatting.Indented);
                    
                    System.Diagnostics.Debug.WriteLine(jsonOutput);
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("detailed_subtype 格式映射关系（JSON）");
                    Console.WriteLine("==========================================");
                    Console.WriteLine(jsonOutput);
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");

                    totalStartTime.Stop();
                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("总耗时统计");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine($"步骤1（文档读取，含格式提取）: {step1StartTime.ElapsedMilliseconds}ms ({step1StartTime.ElapsedMilliseconds / 1000.0:F2}秒)");
                    System.Diagnostics.Debug.WriteLine($"步骤2（按行分类摘要）: {step2StartTime.ElapsedMilliseconds}ms");
                    System.Diagnostics.Debug.WriteLine($"步骤3（控制台输出）: {step3StartTime.ElapsedMilliseconds}ms");
                    System.Diagnostics.Debug.WriteLine($"总耗时: {totalStartTime.ElapsedMilliseconds}ms ({totalStartTime.ElapsedMilliseconds / 1000.0:F2}秒)");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("");

                    // 构建返回数据（简化，主要用于测试）
                    var resultData = new Dictionary<string, object>
                    {
                        { "message", "标题识别完成，结果已输出到控制台" },
                        { "heading_line_count", headingLineCount },
                        { "body_candidate_line_count", classification.BodyLines.Count },
                        { "subtype_format_mapping", subtypeFormatMapping },
                        { "timing", new Dictionary<string, object>
                            {
                                { "step1_document_reading_ms", step1StartTime.ElapsedMilliseconds },
                                { "step2_classify_lines_ms", step2StartTime.ElapsedMilliseconds },
                                { "step3_console_output_ms", step3StartTime.ElapsedMilliseconds },
                                { "total_ms", totalStartTime.ElapsedMilliseconds }
                            }
                        }
                    };

                    return new ToolResult
                    {
                        Success = true,
                        Data = resultData
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[F_test_heading_recognizer] ⚠️ 异常: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"[F_test_heading_recognizer] ⚠️ 堆栈跟踪: {ex.StackTrace}");
                    return new ToolResult { Success = false, Error = $"测试标题识别失败: {ex.Message}" };
                }
            };
        }
    }
}

