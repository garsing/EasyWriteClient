using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 显式开/关 Word 修订跟踪（TrackRevisions）与修订显示（ShowRevisionsAndComments）。
    /// 已下线注册（2026-08）：实现保留，勿再 Register；WS 映射已移除。
    /// </summary>
    public static class F_SetTrackRevisionsTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_set_track_revisions"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序不可用" };
                    }

                    dynamic wordApp = wordApplication;
                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    DocumentState.BindAndActivate(wordApp.ActiveDocument);

                    if (!args.ContainsKey("enabled"))
                    {
                        return new ToolResult { Success = false, Error = "缺少或无效的参数：enabled" };
                    }

                    if (!TryParseBoolRequired(args["enabled"], out bool enabled))
                    {
                        return new ToolResult { Success = false, Error = "缺少或无效的参数：enabled" };
                    }

                    bool showRevisions = enabled;
                    if (args.ContainsKey("show_revisions"))
                    {
                        if (!TryParseBoolRequired(args["show_revisions"], out showRevisions))
                        {
                            return new ToolResult { Success = false, Error = "缺少或无效的参数：show_revisions" };
                        }
                    }

                    Word.Document doc = wordApp.ActiveDocument;

                    try
                    {
                        bool prevTrack = doc.TrackRevisions;
                        bool prevShow = doc.Application.ActiveWindow.View.ShowRevisionsAndComments;

                        doc.TrackRevisions = enabled;
                        doc.Application.ActiveWindow.View.ShowRevisionsAndComments = showRevisions;

                        bool currentTrack = doc.TrackRevisions;
                        bool currentShow = doc.Application.ActiveWindow.View.ShowRevisionsAndComments;

                        System.Diagnostics.Debug.WriteLine(
                            $"[SetTrackRevisions] enabled={enabled}, show={showRevisions}, " +
                            $"prevTrack={prevTrack}, prevShow={prevShow}");

                        string message = enabled
                            ? "已开启修订跟踪（TrackRevisions=true）。"
                            : "已关闭修订跟踪（TrackRevisions=false）；未清除已有未接受修订。";

                        await Task.CompletedTask;

                        return new ToolResult
                        {
                            Success = true,
                            Data = new
                            {
                                track_revisions = currentTrack,
                                show_revisions_and_comments = currentShow,
                                previous_track_revisions = prevTrack,
                                previous_show_revisions_and_comments = prevShow,
                                message = message,
                            },
                        };
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SetTrackRevisions] COM 失败: {ex.Message}");
                        return new ToolResult { Success = false, Error = $"设置审阅模式失败：{ex.Message}" };
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[SetTrackRevisions] 执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"执行失败: {ex.Message}" };
                }
            };
        }

        private static bool TryParseBoolRequired(object value, out bool result)
        {
            result = false;
            if (value == null)
            {
                return false;
            }

            if (value is bool b)
            {
                result = b;
                return true;
            }

            string s = value.ToString()?.Trim().ToLowerInvariant();
            if (s == "true" || s == "1")
            {
                result = true;
                return true;
            }

            if (s == "false" || s == "0")
            {
                result = false;
                return true;
            }

            return false;
        }
    }
}
