using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 列出 Word/WPS 文字稿上的批注。只读，不建检查点。序号是全文序号。
    /// </summary>
    public static class F_ListCommentsTool
    {
        private const string ToolName = "F_list_comments";
        private const string HostHint = "这是 Word/WPS 文字批注，不要用 PPT 或 Excel 工具";
        private const int ScopeTextLimit = 200;

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry[ToolName] = async (args) =>
            {
                try
                {
                    Dictionary<string, object> target = null;
                    if (args != null && args.TryGetValue("target", out object targetRaw) && targetRaw != null)
                    {
                        target = ActionFormatHelper.CoerceDict(targetRaw);
                        if (target == null)
                        {
                            return Fail("target 必须是对象");
                        }

                        InsertTargetLocatorHelper.LocatorParseResult parsed =
                            InsertTargetLocatorHelper.ParseLocatorObject(
                                target, sentenceOnly: false, fieldLabel: "target");
                        if (!parsed.Success)
                        {
                            return Fail("locate_failed: " + parsed.Error);
                        }
                    }

                    if (!DocumentHostAdapter.TryResolveInteropDocument(
                            args,
                            wordApplication,
                            out InteropDocumentHandle docHandle,
                            out ToolResult resolveError))
                    {
                        return WithHostHint(resolveError);
                    }

                    Word.Document doc = docHandle.Document;
                    Word.Range filter = null;
                    if (target != null)
                    {
                        InsertTargetLocatorHelper.LocatorParseResult parsed =
                            InsertTargetLocatorHelper.ParseLocatorObject(
                                target, sentenceOnly: false, fieldLabel: "target");
                        InsertTargetLocatorHelper.LocateRangeKind rangeKind =
                            parsed.Spec.Kind == InsertTargetLocatorHelper.LocatorKind.Sentence
                            || parsed.Spec.Kind == InsertTargetLocatorHelper.LocatorKind.Paragraph
                                ? InsertTargetLocatorHelper.LocateRangeKind.ContentSpan
                                : InsertTargetLocatorHelper.LocateRangeKind.InsertAfter;
                        InsertTargetLocatorHelper.ResolveRangeResult located =
                            InsertTargetLocatorHelper.TryResolveInsertRange(
                                doc,
                                parsed.Spec,
                                rangeKind,
                                docHandle.DisallowBackendApi,
                                refreshDocument: true);
                        if (!located.Success || located.Range == null)
                        {
                            return Fail(string.IsNullOrEmpty(located.Error)
                                ? "定位失败"
                                : located.Error);
                        }

                        filter = located.Range;
                    }

                    var rows = new List<Dictionary<string, object>>();
                    Word.Comments comments = doc.Comments;
                    int count = comments.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Word.Comment comment = comments[i];
                        if (filter != null && !ScopeInside(comment, filter))
                        {
                            continue;
                        }

                        rows.Add(ToRow(i, comment));
                    }

                    string message = rows.Count == 0
                        ? (count == 0 ? "当前稿没有批注。" : "该范围里没有批注。")
                        : "已列出 " + rows.Count + " 条批注。";
                    System.Diagnostics.Debug.WriteLine(
                        "[ListComments] host=" + (docHandle.HostName ?? "")
                        + " total=" + count
                        + " returned=" + rows.Count);

                    await Task.CompletedTask;
                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            ["message"] = message,
                            ["comments"] = rows,
                            ["channel_id"] = docHandle.ChannelId ?? "",
                            ["host"] = docHandle.HostName ?? ""
                        }
                    };
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ListComments] 失败: " + ex.Message);
                    return Fail("列出批注失败: " + ex.Message);
                }
            };
        }

        private static bool ScopeInside(Word.Comment comment, Word.Range filter)
        {
            try
            {
                Word.Range scope = comment.Scope;
                if (scope == null)
                {
                    return false;
                }

                return scope.Start >= filter.Start && scope.End <= filter.End;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ListComments] Scope: " + ex.Message);
                return false;
            }
        }

        private static Dictionary<string, object> ToRow(int index, Word.Comment comment)
        {
            string text = "";
            string author = "";
            string scopeText = "";
            bool point = false;
            try
            {
                text = comment.Range == null ? "" : (comment.Range.Text ?? "");
            }
            catch (Exception)
            {
            }

            try
            {
                author = comment.Author ?? "";
            }
            catch (Exception)
            {
            }

            try
            {
                Word.Range scope = comment.Scope;
                if (scope != null)
                {
                    point = scope.Start == scope.End;
                    scopeText = scope.Text ?? "";
                    if (scopeText.Length > ScopeTextLimit)
                    {
                        scopeText = scopeText.Substring(0, ScopeTextLimit);
                    }
                }
            }
            catch (Exception)
            {
            }

            return new Dictionary<string, object>
            {
                ["index"] = index,
                ["text"] = text,
                ["author"] = author,
                ["scope_text"] = scopeText,
                ["point"] = point,
                ["replies"] = ReadReplies(comment)
            };
        }

        private static List<Dictionary<string, object>> ReadReplies(Word.Comment comment)
        {
            var replies = new List<Dictionary<string, object>>();
            try
            {
                dynamic bag = ((dynamic)comment).Replies;
                int count = (int)bag.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic reply = bag[i];
                    string text = "";
                    string author = "";
                    try
                    {
                        text = reply.Range.Text ?? "";
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        author = reply.Author ?? "";
                    }
                    catch (Exception)
                    {
                    }

                    replies.Add(new Dictionary<string, object>
                    {
                        ["text"] = text,
                        ["author"] = author
                    });
                }
            }
            catch (Exception)
            {
            }

            return replies;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }

        private static ToolResult WithHostHint(ToolResult error)
        {
            if (error == null)
            {
                return Fail(HostHint);
            }

            string text = error.Error ?? "";
            if (text.IndexOf("PPT", StringComparison.Ordinal) < 0)
            {
                error.Error = string.IsNullOrEmpty(text)
                    ? HostHint
                    : text.TrimEnd() + "。" + HostHint;
            }

            return error;
        }
    }
}
