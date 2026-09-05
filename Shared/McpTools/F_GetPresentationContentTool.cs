using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.PresentationHost;

namespace WordAddIn1
{
    public static class F_GetPresentationContentTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_get_presentation_content"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();
                    if (!ChannelContext.TryResolveChannel(args, out IOperationChannel channel, out string resolveError))
                    {
                        return new ToolResult { Success = false, Error = resolveError };
                    }

                    if (!PresentationHostAdapter.TryGetPresentationContent(
                            channel,
                            out PresentationContentResult hostResult,
                            out ToolResult errorResult))
                    {
                        return errorResult;
                    }

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
                        ["channel_id"] = ChannelRegistry.ToPublicId(hostResult.ChannelId) ?? "",
                        ["kind"] = hostResult.Kind ?? "",
                        ["name"] = hostResult.Name ?? "",
                        ["path"] = hostResult.Path ?? "",
                        ["slide_count"] = hostResult.SlideCount,
                        ["display_contents"] = PresentationContentMarkup.BuildDisplayContents(hostResult),
                        ["slides"] = slides
                    };
                    if (hostResult.Truncated)
                    {
                        data["truncated"] = true;
                        if (!string.IsNullOrEmpty(hostResult.TruncatedReason))
                        {
                            data["truncated_reason"] = hostResult.TruncatedReason;
                        }
                    }

                    // 临时关闭：概览不返回 palette。
                    // data["palette"] = hostResult.Palette ?? new List<string>();
                    // data["palette_sampled"] = hostResult.PaletteSampled;
                    // data["palette_scanned_count"] = hostResult.PaletteScannedCount;

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = data
                    };
                }
                catch (OperationCanceledException)
                {
                    return new ToolResult { Success = false, Error = "cancelled by user" };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = "读取演示文稿概览失败: " + ex.Message };
                }
            };
        }
    }
}
