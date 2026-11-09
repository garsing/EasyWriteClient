using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    /// <summary>
    /// 一次导出当前演示文稿里够大的图片。位置写在每页文字里，路径是绝对路径。
    /// </summary>
    public static class F_ExtractPresentationImagesTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_extract_presentation_images"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!PresentationHostAdapter.TryExtractPresentationImages(
                            channel,
                            out PresentationContentResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

                    hostResult.ChannelId = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "";
                    var slides = new List<Dictionary<string, object>>();
                    if (hostResult.Slides != null)
                    {
                        foreach (PresentationSlideInfo slide in hostResult.Slides)
                        {
                            if (slide == null)
                            {
                                continue;
                            }

                            slides.Add(new Dictionary<string, object>
                            {
                                ["index"] = slide.Index,
                                ["slide_id"] = slide.SlideId ?? ""
                            });
                        }
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = hostResult.ChannelId,
                        ["kind"] = hostResult.Kind ?? "",
                        ["name"] = hostResult.Name ?? "",
                        ["path"] = hostResult.Path ?? "",
                        ["slide_count"] = hostResult.SlideCount,
                        ["display_contents"] = PresentationImageExtract.BuildDisplay(hostResult),
                        ["slides"] = slides,
                        ["message"] = "抽取成功"
                    };
                    if (hostResult.Truncated)
                    {
                        data["truncated"] = true;
                        if (!string.IsNullOrEmpty(hostResult.TruncatedReason))
                        {
                            data["truncated_reason"] = hostResult.TruncatedReason;
                        }
                    }

                    ExportImageCaption.ApplyDisplayContents(data);

                    await Task.CompletedTask;
                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "抽取演示文稿图片失败: " + ex.Message };
                }
            };
        }
    }
}
