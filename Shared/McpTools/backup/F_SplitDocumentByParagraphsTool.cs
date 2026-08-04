using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按段落划分文档内容工具
    /// 将Word文档按段落划分，并以JSON格式返回每个段落的内容
    /// </summary>
    public static class F_SplitDocumentByParagraphsTool
    {
        /// <summary>
        /// 注册按段落划分文档内容工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            List<McpTool> availableTools,
            object wordApplication)
        {
            var tool = new McpTool
            {
                name = "F_split_document_by_paragraphs",
                alias = "按段落划分文档",
                description = "按段落划分Word文档内容，返回包含段落序号和内容的JSON数组。注意：此工具从最新快照获取内容，而不是直接从Word文档读取，因为Word文档可能显示hunk形式（包含待删除和新增的内容），而最新快照才是真正修改后的文章内容。",
                usage = "当需要将文档内容按段落进行划分和展示时使用此工具。返回的是最新快照中的内容，即用户确认所有修改后的最终文档内容。",
                input_schema = new McpTool.ToolInputSchema
                {
                    Type = "object",
                    Properties = new Dictionary<string, object>(),
                    required = new List<string>() // 没有必需参数
                }
            };

            toolRegistry["F_split_document_by_paragraphs"] = async (args) =>
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

                    var paragraphs = GetParagraphsContentFromSnapshot(wordApp.ActiveDocument);

                    // 打印到调试控制台
                    System.Diagnostics.Debug.WriteLine("=== 文档段落内容 ===");
                    System.Diagnostics.Debug.WriteLine($"文档共有 {paragraphs.Count} 个段落");
                    foreach (dynamic para in paragraphs)
                    {
                        System.Diagnostics.Debug.WriteLine($"段落 {para.index}: {para.content}");
                    }
                    System.Diagnostics.Debug.WriteLine("======================");

                    await Task.CompletedTask; // 确保异步执行

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            document_name = wordApp.ActiveDocument.Name ?? "未命名文档",
                            total_paragraphs = paragraphs.Count,
                            paragraphs = paragraphs
                        }
                    };

                    return result;
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"按段落划分文档内容失败: {ex.Message}" };
                }
            };

            availableTools.Add(tool);
        }

        /// <summary>
        /// 从最新快照获取文档中所有段落的内容
        /// 因为Word文档可能显示hunk形式（包含待删除和新增的内容），
        /// 而最新快照才是真正修改后的文章内容
        /// 输出原汁原味的文档段落内容，不做任何过滤处理
        /// </summary>
        private static List<object> GetParagraphsContentFromSnapshot(Word.Document doc)
        {
            var paragraphs = new List<object>();

            try
            {
                // 获取最新快照
                var currentSnapshot = DocumentState.CurrentSnapshot;
                
                if (currentSnapshot == null || currentSnapshot.Count == 0)
                {
                    // 没有快照，需要初始化
                    System.Diagnostics.Debug.WriteLine("[DEBUG] 没有可用的快照，初始化段落跟踪");
                    
                    // 从Word文档读取段落内容（包含所有段落，包括空段落）
                    var paragraphsBefore = new List<string>();
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        var para = doc.Paragraphs[i];
                        string content = para.Range.Text?.TrimEnd('\r', '\n') ?? "";
                        // 包含所有段落，包括空段落
                        paragraphsBefore.Add(content);
                    }
                    
                    // 初始化段落跟踪
                    DocumentState.InitializeParagraphTracking(paragraphsBefore);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 段落跟踪初始化完成，共 {paragraphsBefore.Count} 个段落（包含空段落）");
                    
                    // 重新获取快照
                    currentSnapshot = DocumentState.CurrentSnapshot;
                }

                // 遍历快照中的每个段落名称，获取对应的内容
                for (int i = 0; i < currentSnapshot.Count; i++)
                {
                    string paragraphName = currentSnapshot[i];
                    string content = DocumentState.GetParagraphContent(paragraphName) ?? "";

                    // 不做任何过滤处理，输出原汁原味的内容
                    paragraphs.Add(new
                    {
                        index = i + 1, // 使用快照中的索引作为段落序号
                        content = content
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 从快照获取段落内容失败: {ex.Message}");
            }

            return paragraphs;
        }

    }
}
