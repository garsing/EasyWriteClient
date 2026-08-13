using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 获取Word文档结构信息工具（已下线，不再 Register；请用 F_get_document_content）
    /// 代码已按 DocumentHost 迁入，便于日后重新启用；只读不抢前台。
    /// </summary>
    public static class F_GetDocumentStructureTool
    {
        /// <summary>
        /// 注册获取文档结构工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_document_structure"] = async (args) =>
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] get_document_structure工具开始执行");

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
                        $"[F_get_document_structure] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    bool includeContent = args.ContainsKey("include_content") && Convert.ToBoolean(args["include_content"]);
                    int maxPreviewLength = args.ContainsKey("max_preview_length") ? Convert.ToInt32(args["max_preview_length"]) : 50;

                    var documentStructure = new
                    {
                        document_name = document.Name ?? "未命名文档",
                        paragraphs = GetParagraphsInfo(document, includeContent, maxPreviewLength),
                        tables = GetTablesInfo(document, includeContent, maxPreviewLength),
                        images = GetImagesInfo(document, includeContent, maxPreviewLength)
                    };

                    await Task.CompletedTask; // 避免异步方法警告

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = documentStructure
                    };

                    System.Diagnostics.Debug.WriteLine("[DEBUG] get_document_structure工具执行成功");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] get_document_structure工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"获取文档结构失败: {ex.Message}" };
                }
            };
        }

        /// <summary>
        /// 获取文档中所有段落的信息
        /// </summary>
        private static List<object> GetParagraphsInfo(Word.Document doc, bool includeContent, int maxPreviewLength)
        {
            var paragraphs = new List<object>();

            try
            {
                for (int i = 1; i <= doc.Paragraphs.Count; i++)
                {
                    Word.Paragraph para = doc.Paragraphs[i];
                    string content = includeContent ? para.Range.Text?.Trim() : "";
                    string preview = content.Length > maxPreviewLength
                        ? content.Substring(0, maxPreviewLength) + "..."
                        : content;

                    paragraphs.Add(new
                    {
                        index = i,
                        start_position = para.Range.Start,
                        end_position = para.Range.End,
                        length = para.Range.End - para.Range.Start,
                        preview = preview,
                        has_content = !string.IsNullOrEmpty(content)
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取段落信息失败: {ex.Message}");
            }

            return paragraphs;
        }

        /// <summary>
        /// 获取文档中所有表格的信息
        /// </summary>
        private static List<object> GetTablesInfo(Word.Document doc, bool includeContent, int maxPreviewLength)
        {
            var tables = new List<object>();

            try
            {
                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    Word.Table table = doc.Tables[i];
                    string content = includeContent ? table.Range.Text?.Trim() : "";
                    string preview = content.Length > maxPreviewLength
                        ? content.Substring(0, maxPreviewLength) + "..."
                        : content;

                    tables.Add(new
                    {
                        index = i,
                        start_position = table.Range.Start,
                        end_position = table.Range.End,
                        length = table.Range.End - table.Range.Start,
                        rows = table.Rows.Count,
                        columns = table.Columns.Count,
                        preview = preview,
                        has_content = !string.IsNullOrEmpty(content)
                    });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取表格信息失败: {ex.Message}");
            }

            return tables;
        }

        /// <summary>
        /// 获取文档中所有图片的信息
        /// </summary>
        private static List<object> GetImagesInfo(Word.Document doc, bool includeContent, int maxPreviewLength)
        {
            var images = new List<object>();
            int imageCount = 0;

            try
            {
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (shape.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                    {
                        imageCount++;
                        string content = includeContent ? $"图片 {imageCount}" : "";
                        string preview = content.Length > maxPreviewLength
                            ? content.Substring(0, maxPreviewLength) + "..."
                            : content;

                        images.Add(new
                        {
                            index = imageCount,
                            start_position = shape.Range.Start,
                            end_position = shape.Range.End,
                            length = shape.Range.End - shape.Range.Start,
                            width = shape.Width,
                            height = shape.Height,
                            preview = preview,
                            has_content = !string.IsNullOrEmpty(content)
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 获取图片信息失败: {ex.Message}");
            }

            return images;
        }
    }
}
