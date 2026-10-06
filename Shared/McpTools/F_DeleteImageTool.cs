using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 按 image_id 删除普通 InlineShape 图片（I_，非 Word Chart）。经 <see cref="DocumentHostAdapter"/>（Word/WPS）。
    /// </summary>
    public static class F_DeleteImageTool
    {
        private const string ChartHint = "这是普通图片编号，请用 F_delete_image。Word Chart 才用 F_delete_chart。";
        private const string TableHint = "这是表格编号，请用 F_delete_table。";
        private const string ProcessHint = "不要用 F_process_document_actions 删除图片。";

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_delete_image"] = async (args) =>
            {
                try
                {
                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return WithProcessHint(resolveError);
                    }

                    Word.Document document = docHandle.Document;
                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteImage] host={docHandle.HostName}, channel_id={docHandle.ChannelId}");

                    string imageId = args != null && args.ContainsKey("image_id")
                        ? args["image_id"]?.ToString()?.Trim()
                        : "";
                    string idErr = ValidateImageId(imageId);
                    if (idErr != null)
                    {
                        return Fail(idErr);
                    }

                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        return Fail($"生成图片映射失败: {ex.Message}");
                    }

                    Word.InlineShape shape = ImageResolveHelper.ResolveInlineShapeByImageId(document, imageId);
                    if (shape == null)
                    {
                        return Fail($"找不到图片编号 '{imageId}'，请确认是否为当前稿的 I_ 编号");
                    }

                    int startPos = TryGetRangeStart(shape);
                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteImage] 即将删除 image_id={imageId}, startPos={startPos}");

                    try
                    {
                        shape.Delete();
                    }
                    catch (Exception ex)
                    {
                        return Fail($"删除图片失败: {ex.Message}");
                    }

                    bool mappingRefreshed = true;
                    try
                    {
                        WordReader.ReadWord(document);
                    }
                    catch (Exception ex)
                    {
                        mappingRefreshed = false;
                        System.Diagnostics.Debug.WriteLine(
                            $"[DeleteImage] 删后刷新映射失败（图已删，不回滚）: {ex.Message}");
                    }

                    bool cursorMoved = TryNavigateToStart(
                        wordApplication as Word.Application, document, startPos, imageId);

                    string message = BuildSuccessMessage(imageId, cursorMoved, mappingRefreshed);
                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteImage] 成功 image_id={imageId}, startPos={startPos}, " +
                        $"cursor_moved={cursorMoved}, mapping_refreshed={mappingRefreshed}");

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["image_deleted"] = true,
                            ["image_id"] = imageId,
                            ["channel_id"] = docHandle.ChannelId ?? "",
                            ["host"] = docHandle.HostName ?? "",
                            ["message"] = message
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DeleteImage] 失败: {ex.Message}");
                    return Fail($"删除图片失败: {ex.Message}");
                }
            };
        }

        private static string ValidateImageId(string imageId)
        {
            if (string.IsNullOrWhiteSpace(imageId))
            {
                return "未提供图片编号";
            }

            if (imageId.StartsWith("C_", StringComparison.OrdinalIgnoreCase))
            {
                return ChartHint;
            }

            if (imageId.StartsWith("T_", StringComparison.OrdinalIgnoreCase))
            {
                return TableHint;
            }

            if (!imageId.StartsWith("I_", StringComparison.OrdinalIgnoreCase))
            {
                return "image_id 须为 I_ 前缀的图片编号。" + ProcessHint;
            }

            return null;
        }

        private static int TryGetRangeStart(Word.InlineShape shape)
        {
            try
            {
                return shape.Range.Start;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DeleteImage] 读取 Range.Start 失败: {ex.Message}");
                return -1;
            }
        }

        private static bool TryNavigateToStart(
            Word.Application app,
            Word.Document document,
            int startPos,
            string imageId)
        {
            if (app == null || document == null)
            {
                return false;
            }

            try
            {
                int maxPos = Math.Max(0, document.Content.End - 1);
                int pos = startPos < 0 ? maxPos : Math.Min(Math.Max(0, startPos), maxPos);
                Word.Range atStart = document.Range(pos, pos);
                return PostModifyNavigateHelper.NavigateAfterEnd(
                    app, atStart, $"delete_image:{imageId}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[DeleteImage] 移光标失败（图已删，不回滚）: {ex.Message}");
                return false;
            }
        }

        private static string BuildSuccessMessage(string imageId, bool cursorMoved, bool mappingRefreshed)
        {
            var sb = new StringBuilder();
            sb.Append($"图片 {imageId} 已删除。");
            if (cursorMoved)
            {
                sb.Append("光标已停留在原图位置。");
            }
            else
            {
                sb.Append("未能将光标移到原图位置。");
            }

            sb.Append("重插请再调用 F_insert_image 并带 target，不要省略 target。");
            sb.Append("不要用 F_process_document_actions 删图，不要对 I_ 调用 F_delete_chart。");
            if (!mappingRefreshed)
            {
                sb.Append("删后映射刷新失败，建议调用 F_get_document_content。");
            }

            return sb.ToString();
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }

        private static ToolResult WithProcessHint(ToolResult error)
        {
            if (error == null)
            {
                return Fail(ProcessHint);
            }

            string text = error.Error ?? "";
            if (text.IndexOf("F_process_document_actions", StringComparison.Ordinal) < 0)
            {
                error.Error = string.IsNullOrEmpty(text)
                    ? ProcessHint
                    : text.TrimEnd() + "。" + ProcessHint;
            }

            return error;
        }
    }
}
