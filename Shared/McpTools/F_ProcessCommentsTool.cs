using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using WordAddIn1.DocumentHost;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 在 Word/WPS 文字稿上增加或删除批注。增加走统一定位；删除用调用开始时的全文序号。
    /// </summary>
    public static class F_ProcessCommentsTool
    {
        private const string ToolName = "F_process_comments";
        private const string HostHint = "这是 Word/WPS 文字批注，不要用 PPT 或 Excel 工具";

        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry[ToolName] = async (args) =>
            {
                if (TryGetContractParseFailure(args, out ToolResult contractFail))
                {
                    return contractFail;
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
                var items = ReadItems(args["comments"]);
                var snapshot = SnapshotComments(doc);
                int pinStart;
                int pinEnd;
                try
                {
                    Word.Selection selection = doc.Application.Selection;
                    pinStart = selection.Start;
                    pinEnd = selection.End;
                }
                catch (Exception ex)
                {
                    pinStart = -1;
                    pinEnd = -1;
                    System.Diagnostics.Debug.WriteLine(
                        "[ProcessComments] 读选区失败: " + ex.Message);
                }

                System.Diagnostics.Debug.WriteLine(
                    "[ProcessComments] host=" + (docHandle.HostName ?? "")
                    + " channel_id=" + (docHandle.ChannelId ?? "")
                    + " items=" + items.Count
                    + " comments=" + snapshot.Count);

                var results = new List<Dictionary<string, object>>();
                var deleted = new HashSet<int>();
                int addCount = 0;
                int deleteCount = 0;
                bool anyFailed = false;
                int lastAddStart = -1;
                int lastAddEnd = -1;
                bool haveLastAdd = false;

                for (int i = 0; i < items.Count; i++)
                {
                    CommentItem item = items[i];
                    var row = new Dictionary<string, object>
                    {
                        ["index"] = i,
                        ["type"] = item.Type,
                        ["success"] = false
                    };
                    try
                    {
                        if (item.Type == "add")
                        {
                            if (!TryAdd(
                                    doc,
                                    docHandle.DisallowBackendApi,
                                    item,
                                    pinStart,
                                    pinEnd,
                                    out string anchor,
                                    out int rangeStart,
                                    out int rangeEnd,
                                    out string error))
                            {
                                anyFailed = true;
                                row["error"] = error;
                            }
                            else
                            {
                                addCount++;
                                row["success"] = true;
                                row["anchor"] = anchor;
                                lastAddStart = rangeStart;
                                lastAddEnd = rangeEnd;
                                haveLastAdd = true;
                            }
                        }
                        else if (!TryDelete(snapshot, deleted, item.CommentIndex, out string error))
                        {
                            anyFailed = true;
                            row["error"] = error;
                            row["comment_index"] = item.CommentIndex;
                        }
                        else
                        {
                            deleteCount++;
                            row["success"] = true;
                            row["comment_index"] = item.CommentIndex;
                        }
                    }
                    catch (Exception ex)
                    {
                        anyFailed = true;
                        row["error"] = ex.Message;
                        System.Diagnostics.Debug.WriteLine(
                            "[ProcessComments] item " + i + " " + ex.Message);
                    }

                    results.Add(row);
                }

                if (addCount > 0)
                {
                    TryShowComments(doc);
                    TrySelect(doc, lastAddStart, lastAddEnd, haveLastAdd);
                }

                bool mutated = addCount + deleteCount > 0;
                int total = items.Count;
                var data = new Dictionary<string, object>
                {
                    ["message"] = anyFailed
                        ? "有批注未处理完成。"
                        : "已处理 " + total + " 条批注。",
                    ["add_count"] = addCount,
                    ["delete_count"] = deleteCount,
                    ["mutated"] = mutated,
                    ["results"] = results,
                    ["channel_id"] = docHandle.ChannelId ?? "",
                    ["host"] = docHandle.HostName ?? ""
                };
                if (!mutated)
                {
                    data["checkpoint_cleanup_only"] = true;
                }

                System.Diagnostics.Debug.WriteLine(
                    "[ProcessComments] done add=" + addCount
                    + " delete=" + deleteCount
                    + " failed=" + anyFailed);

                await Task.CompletedTask;
                if (anyFailed)
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = "有批注未处理完成。",
                        Data = data
                    };
                }

                return new ToolResult { Success = true, Data = data };
            };
        }

        public static bool TryGetContractParseFailure(
            Dictionary<string, object> args,
            out ToolResult failure)
        {
            failure = null;
            if (args == null || !args.ContainsKey("comments"))
            {
                failure = DocumentCheckpointMutatingTools.Fail("缺少必需参数：comments");
                return true;
            }

            if (!TryEnumerateItems(args["comments"], out List<object> rawItems, out string error))
            {
                failure = DocumentCheckpointMutatingTools.Fail(error);
                return true;
            }

            if (rawItems.Count == 0)
            {
                failure = DocumentCheckpointMutatingTools.Fail("comments 不能为空");
                return true;
            }

            foreach (object raw in rawItems)
            {
                Dictionary<string, object> dict = ActionFormatHelper.CoerceDict(raw);
                if (dict == null)
                {
                    failure = DocumentCheckpointMutatingTools.Fail("comments 的每一条都必须是对象");
                    return true;
                }

                if (!dict.TryGetValue("type", out object typeRaw) || typeRaw == null)
                {
                    failure = DocumentCheckpointMutatingTools.Fail("comments 的每一条都必须有 type");
                    return true;
                }

                string type = typeRaw.ToString().Trim().ToLowerInvariant();
                if (type != "add" && type != "delete")
                {
                    failure = DocumentCheckpointMutatingTools.Fail(
                        "type 只能是 add 或 delete");
                    return true;
                }

                if (dict.ContainsKey("author"))
                {
                    failure = DocumentCheckpointMutatingTools.Fail("不接受 author，作者用当前 Word/WPS 用户");
                    return true;
                }

                if (type == "add")
                {
                    if (dict.ContainsKey("index"))
                    {
                        failure = DocumentCheckpointMutatingTools.Fail("add 不接受 index");
                        return true;
                    }

                    if (!dict.TryGetValue("text", out object textRaw) || textRaw == null
                        || string.IsNullOrWhiteSpace(textRaw.ToString()))
                    {
                        failure = DocumentCheckpointMutatingTools.Fail("add 的 text 不能为空");
                        return true;
                    }

                    if (dict.TryGetValue("target", out object targetRaw) && targetRaw != null)
                    {
                        Dictionary<string, object> target = ActionFormatHelper.CoerceDict(targetRaw);
                        if (target == null)
                        {
                            failure = DocumentCheckpointMutatingTools.Fail("add 的 target 必须是对象");
                            return true;
                        }

                        InsertTargetLocatorHelper.LocatorParseResult parsed =
                            InsertTargetLocatorHelper.ParseLocatorObject(
                                target, sentenceOnly: false, fieldLabel: "target");
                        if (!parsed.Success)
                        {
                            failure = DocumentCheckpointMutatingTools.Fail(
                                "locate_failed: " + parsed.Error);
                            return true;
                        }
                    }
                }
                else
                {
                    if (dict.ContainsKey("text") || dict.ContainsKey("target"))
                    {
                        failure = DocumentCheckpointMutatingTools.Fail(
                            "delete 只接受 index，不接受 text 或 target");
                        return true;
                    }

                    if (!dict.TryGetValue("index", out object indexRaw)
                        || !TryReadIndex(indexRaw, out _))
                    {
                        failure = DocumentCheckpointMutatingTools.Fail(
                            "delete 的 index 必须是从 1 开始的整数");
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool TryAdd(
            Word.Document doc,
            bool disallowBackendApi,
            CommentItem item,
            int pinStart,
            int pinEnd,
            out string anchor,
            out int rangeStart,
            out int rangeEnd,
            out string error)
        {
            anchor = null;
            rangeStart = 0;
            rangeEnd = 0;
            error = null;
            Word.Range range;
            if (item.Target == null)
            {
                if (pinStart < 0)
                {
                    error = "没有活动选区";
                    return false;
                }

                range = doc.Range(pinStart, pinEnd);
                anchor = pinStart == pinEnd ? "cursor" : "selection";
            }
            else
            {
                InsertTargetLocatorHelper.LocatorParseResult parsed =
                    InsertTargetLocatorHelper.ParseLocatorObject(
                        item.Target, sentenceOnly: false, fieldLabel: "target");
                if (!parsed.Success)
                {
                    error = parsed.Error;
                    return false;
                }

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
                        disallowBackendApi,
                        refreshDocument: true);
                if (!located.Success || located.Range == null)
                {
                    error = string.IsNullOrEmpty(located.Error) ? "定位失败" : located.Error;
                    return false;
                }

                range = located.Range;
                anchor = AnchorName(parsed.Spec.Kind);
            }

            rangeStart = range.Start;
            rangeEnd = range.End;
            object text = item.Text;
            doc.Comments.Add(range, ref text);
            return true;
        }

        private static bool TryDelete(
            List<Word.Comment> snapshot,
            HashSet<int> deleted,
            int index,
            out string error)
        {
            error = null;
            if (index < 1 || index > snapshot.Count)
            {
                error = "序号超出开始时的批注条数";
                return false;
            }

            if (!deleted.Add(index))
            {
                error = "同一序号已经删过";
                return false;
            }

            snapshot[index - 1].Delete();
            return true;
        }

        private static List<Word.Comment> SnapshotComments(Word.Document doc)
        {
            var list = new List<Word.Comment>();
            Word.Comments comments = doc.Comments;
            int count = comments.Count;
            for (int i = 1; i <= count; i++)
            {
                list.Add(comments[i]);
            }

            return list;
        }

        private static void TryShowComments(Word.Document doc)
        {
            try
            {
                dynamic view = doc.ActiveWindow.View;
                view.ShowComments = true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ProcessComments] ShowComments: " + ex.Message);
            }
        }

        private static void TrySelect(Word.Document doc, int start, int end, bool have)
        {
            if (!have)
            {
                return;
            }

            try
            {
                doc.Application.Selection.SetRange(start, end);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ProcessComments] 移光标失败: " + ex.Message);
            }
        }

        private static string AnchorName(InsertTargetLocatorHelper.LocatorKind kind)
        {
            switch (kind)
            {
                case InsertTargetLocatorHelper.LocatorKind.Sentence:
                    return "sentence";
                case InsertTargetLocatorHelper.LocatorKind.Paragraph:
                    return "paragraph";
                case InsertTargetLocatorHelper.LocatorKind.EmptyCell:
                    return "empty_cell";
                case InsertTargetLocatorHelper.LocatorKind.Table:
                    return "table";
                case InsertTargetLocatorHelper.LocatorKind.Chart:
                    return "chart";
                case InsertTargetLocatorHelper.LocatorKind.Image:
                    return "image";
                case InsertTargetLocatorHelper.LocatorKind.DocumentStart:
                    return "document_start";
                default:
                    return "sentence";
            }
        }

        private static List<CommentItem> ReadItems(object raw)
        {
            TryEnumerateItems(raw, out List<object> rawItems, out _);
            var items = new List<CommentItem>();
            foreach (object one in rawItems)
            {
                Dictionary<string, object> dict = ActionFormatHelper.CoerceDict(one);
                string type = dict["type"].ToString().Trim().ToLowerInvariant();
                var item = new CommentItem { Type = type };
                if (type == "add")
                {
                    item.Text = dict["text"].ToString().Trim();
                    if (dict.TryGetValue("target", out object targetRaw) && targetRaw != null)
                    {
                        item.Target = ActionFormatHelper.CoerceDict(targetRaw);
                    }
                }
                else
                {
                    TryReadIndex(dict["index"], out int index);
                    item.CommentIndex = index;
                }

                items.Add(item);
            }

            return items;
        }

        private static bool TryEnumerateItems(
            object raw,
            out List<object> items,
            out string error)
        {
            items = new List<object>();
            error = null;
            if (raw == null || raw is string)
            {
                error = "comments 必须是数组";
                return false;
            }

            if (!(raw is IEnumerable enumerable))
            {
                error = "comments 必须是数组";
                return false;
            }

            foreach (object one in enumerable)
            {
                items.Add(one);
            }

            return true;
        }

        private static bool TryReadIndex(object raw, out int index)
        {
            index = 0;
            if (raw == null || raw is bool)
            {
                return false;
            }

            double number;
            try
            {
                number = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return false;
            }

            if (number < 1 || Math.Abs(number - Math.Round(number)) > 0.0000001)
            {
                return false;
            }

            if (number > int.MaxValue)
            {
                return false;
            }

            index = (int)Math.Round(number);
            return true;
        }

        private static ToolResult WithHostHint(ToolResult error)
        {
            if (error == null)
            {
                return new ToolResult { Success = false, Error = HostHint };
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

        private sealed class CommentItem
        {
            public string Type;
            public string Text;
            public Dictionary<string, object> Target;
            public int CommentIndex;
        }
    }
}
