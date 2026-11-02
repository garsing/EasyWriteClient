using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1.BrowserHost
{
    internal sealed class BrowserInteractStep
    {
        public string Action { get; set; }

        public string Ref { get; set; }

        public string Text { get; set; }

        public string Direction { get; set; }

        public string Key { get; set; }

        public string Option { get; set; }
    }

    internal sealed class BrowserInteractItemResult
    {
        public int ActionIndex { get; set; }

        public string Action { get; set; }

        public string Ref { get; set; }

        public string Status { get; set; }

        public string Error { get; set; }

        public Dictionary<string, object> ToResponseObject()
        {
            var row = new Dictionary<string, object>
            {
                ["action_index"] = ActionIndex,
                ["action"] = Action ?? "",
                ["status"] = Status ?? "failed",
            };
            if (!string.IsNullOrEmpty(Ref))
            {
                row["ref"] = Ref;
            }

            if (!string.IsNullOrEmpty(Error))
            {
                row["error"] = Error;
            }

            return row;
        }
    }

    internal static class BrowserInteractBatch
    {
        internal static readonly HashSet<string> AllowedKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Enter", "Return", "Tab", "Escape", "Esc", "Backspace", "Delete",
            "ArrowUp", "ArrowDown", "ArrowLeft", "ArrowRight",
            "Home", "End", "PageUp", "PageDown"
        };

        public static bool TryCompile(
            Dictionary<string, object> args,
            out List<BrowserInteractStep> steps,
            out string contractError)
        {
            steps = new List<BrowserInteractStep>();
            contractError = null;
            if (args != null && args.ContainsKey("actions") && args["actions"] != null)
            {
                if (!TryReadObjectList(args["actions"], out List<Dictionary<string, object>> rows, out contractError))
                {
                    return false;
                }

                if (rows.Count == 0)
                {
                    contractError = "actions 不能为空";
                    return false;
                }

                foreach (Dictionary<string, object> row in rows)
                {
                    steps.Add(ReadStep(row));
                }

                return true;
            }

            string action = ReadString(args, "action");
            if (string.IsNullOrWhiteSpace(action))
            {
                contractError = "必须提供 action（click|hover|type|scroll|press|select）或 actions";
                return false;
            }

            steps.Add(ReadStep(args));
            return true;
        }

        public static async Task<ToolResult> Run(BrowserChannel channel, List<BrowserInteractStep> steps)
        {
            var items = new List<BrowserInteractItemResult>();
            BrowserInteractResult lastSuccess = null;
            string stopReason = null;
            bool stopped = false;

            for (int i = 0; i < steps.Count; i++)
            {
                BrowserInteractStep step = steps[i] ?? new BrowserInteractStep();
                if (stopped)
                {
                    items.Add(Skipped(i, step, stopReason));
                    continue;
                }

                string itemError = ValidateItem(step);
                if (itemError != null)
                {
                    items.Add(Failed(i, step, itemError));
                    continue;
                }

                BrowserInteractResult performed;
                try
                {
                    performed = await BrowserHostAdapter
                        .InteractStepAsync(
                            channel,
                            step.Action,
                            step.Ref,
                            step.Text,
                            step.Direction,
                            step.Key,
                            step.Option)
                        .ConfigureAwait(true);
                }
                catch (Exception ex)
                {
                    items.Add(Failed(i, step, "F_browser_interact 失败: " + ex.Message));
                    continue;
                }

                if (performed == null || !performed.Success)
                {
                    items.Add(Failed(i, step, performed != null ? performed.Error : "interact 失败"));
                    continue;
                }

                items.Add(Succeeded(i, step, performed.Ref));
                lastSuccess = performed;
                if (performed.StopBatch)
                {
                    stopped = true;
                    stopReason = performed.StopReason
                        ?? "前一项改变了页面，后面的项未执行";
                }
            }

            return await FinishAsync(channel, items, lastSuccess).ConfigureAwait(true);
        }

        public static string StripHash(string url)
        {
            if (string.IsNullOrEmpty(url))
            {
                return "";
            }

            int hash = url.IndexOf('#');
            return hash < 0 ? url : url.Substring(0, hash);
        }

        public static bool UrlChanged(string before, string after)
        {
            return !string.Equals(StripHash(before), StripHash(after), StringComparison.OrdinalIgnoreCase);
        }

        public static async Task<BrowserInteractResult> MaybeCompleteAsync(
            BrowserInteractResult performed,
            bool completeAfter)
        {
            if (performed == null || !performed.Success)
            {
                return performed;
            }

            if (performed.ExpandedFrame)
            {
                return performed;
            }

            if (!completeAfter)
            {
                return performed;
            }

            return await BrowserResnapshotAfterAction
                .CompleteInteractAsync(
                    performed.Channel,
                    performed.Action,
                    performed.Ref,
                    performed.Layer ?? BrowserResnapshotLayer.Shell(),
                    performed.Message,
                    performed.OpenedNewTab,
                    performed.FromChannelId)
                .ConfigureAwait(true);
        }

        internal static string BuildMessage(List<BrowserInteractItemResult> items)
        {
            int succeeded = 0;
            foreach (BrowserInteractItemResult item in items)
            {
                if (item != null && item.Status == "success")
                {
                    succeeded++;
                }
            }

            if (succeeded == items.Count && items.Count > 0)
            {
                return "全部成功。";
            }

            var sb = new StringBuilder();
            if (succeeded == 0)
            {
                sb.Append("全部失败。");
            }
            else
            {
                sb.Append("部分成功。已做成且不要重做：");
                bool first = true;
                foreach (BrowserInteractItemResult item in items)
                {
                    if (item == null || item.Status != "success")
                    {
                        continue;
                    }

                    if (!first)
                    {
                        sb.Append("、");
                    }

                    first = false;
                    sb.Append("action[").Append(item.ActionIndex).Append("] ").Append(item.Action ?? "");
                }

                sb.Append("。失败：");
            }

            bool anyFail = false;
            foreach (BrowserInteractItemResult item in items)
            {
                if (item == null || item.Status == "success")
                {
                    continue;
                }

                if (anyFail)
                {
                    sb.Append("；");
                }

                anyFail = true;
                string label = item.Status == "skipped" ? "未执行" : "失败";
                sb.Append(label).Append("：action[").Append(item.ActionIndex).Append("] ").Append(item.Action ?? "");
                if (!string.IsNullOrEmpty(item.Error))
                {
                    sb.Append("，").Append(item.Error);
                }
            }

            return sb.ToString();
        }

        private static async Task<ToolResult> FinishAsync(
            BrowserChannel fallbackChannel,
            List<BrowserInteractItemResult> items,
            BrowserInteractResult lastSuccess)
        {
            int succeeded = 0;
            foreach (BrowserInteractItemResult item in items)
            {
                if (item != null && item.Status == "success")
                {
                    succeeded++;
                }
            }

            bool allOk = items.Count > 0 && succeeded == items.Count;
            string outcome = allOk
                ? "all_succeeded"
                : (succeeded > 0 ? "partial_succeeded" : "all_failed");
            string message = BuildMessage(items);

            BrowserInteractResult snap = null;
            if (succeeded > 0 && lastSuccess != null)
            {
                if (lastSuccess.ExpandedFrame && !string.IsNullOrWhiteSpace(lastSuccess.Snapshot))
                {
                    snap = lastSuccess;
                }
                else
                {
                    snap = await BrowserResnapshotAfterAction
                        .CompleteInteractAsync(
                            lastSuccess.Channel ?? fallbackChannel,
                            lastSuccess.Action,
                            lastSuccess.Ref,
                            lastSuccess.Layer ?? BrowserResnapshotLayer.Shell(),
                            lastSuccess.Message,
                            lastSuccess.OpenedNewTab,
                            lastSuccess.FromChannelId)
                        .ConfigureAwait(true);
                }
            }

            BrowserChannel outChannel = snap != null && snap.Channel != null
                ? snap.Channel
                : (lastSuccess != null && lastSuccess.Channel != null ? lastSuccess.Channel : fallbackChannel);

            var data = new Dictionary<string, object>
            {
                ["channel_id"] = outChannel != null ? ChannelRegistry.ToPublicId(outChannel.ChannelId) : "",
                ["kind"] = "browser",
                ["track"] = outChannel != null ? outChannel.Track : "",
                ["batch_outcome"] = outcome,
                ["message"] = message,
                ["results"] = items.ConvertAll(r => (object)r.ToResponseObject()),
            };

            BrowserInteractItemResult lastItem = items.Count > 0 ? items[items.Count - 1] : null;
            if (lastSuccess != null)
            {
                data["action"] = lastSuccess.Action ?? "";
                if (!string.IsNullOrEmpty(lastSuccess.Ref))
                {
                    data["ref"] = lastSuccess.Ref;
                }

                data["url"] = snap != null ? (snap.Url ?? "") : (lastSuccess.Url ?? "");
                data["title"] = snap != null ? (snap.Title ?? "") : (lastSuccess.Title ?? "");
                data["refs_invalidated"] = snap != null && snap.RefsInvalidated;
                if (lastSuccess.OpenedNewTab && lastSuccess.Channel != null)
                {
                    data["opened_tab"] = true;
                    data["opened_channel_id"] = ChannelRegistry.ToPublicId(lastSuccess.Channel.ChannelId);
                    if (!string.IsNullOrWhiteSpace(lastSuccess.FromChannelId))
                    {
                        data["from_channel_id"] = ChannelRegistry.ToPublicId(lastSuccess.FromChannelId);
                    }
                }
            }
            else if (lastItem != null)
            {
                data["action"] = lastItem.Action ?? "";
                if (!string.IsNullOrEmpty(lastItem.Ref))
                {
                    data["ref"] = lastItem.Ref;
                }

                data["url"] = fallbackChannel != null ? (fallbackChannel.Url ?? "") : "";
                data["title"] = fallbackChannel != null ? (fallbackChannel.Title ?? "") : "";
                data["refs_invalidated"] = false;
            }

            if (allOk)
            {
                data["message"] = "全部成功。";
                if (snap != null && !string.IsNullOrWhiteSpace(snap.Message)
                    && snap.Message.IndexOf("新 ref", StringComparison.Ordinal) >= 0)
                {
                    data["message"] = "全部成功。请使用本次返回的新 ref。";
                }
            }

            if (snap != null)
            {
                BrowserResnapshotAfterAction.WriteSnapshotFields(
                    data,
                    snap.ExpandedFrame,
                    snap.Resnapshot || !snap.ExpandedFrame,
                    snap.Snapshot,
                    snap.Mode,
                    snap.Truncated,
                    snap.TruncatedReason);
            }

            return new ToolResult
            {
                Success = allOk,
                Error = allOk ? null : message,
                Data = data
            };
        }

        private static string ValidateItem(BrowserInteractStep step)
        {
            string act = (step.Action ?? "").Trim().ToLowerInvariant();
            if (act != "click" && act != "hover" && act != "type" && act != "scroll"
                && act != "press" && act != "select")
            {
                return "action 须为 click|hover|type|scroll|press|select";
            }

            if (act == "press")
            {
                if (string.IsNullOrWhiteSpace(step.Key))
                {
                    return "必须提供 key";
                }

                if (!AllowedKeys.Contains(step.Key.Trim()))
                {
                    return "key 不在白名单: " + step.Key;
                }
            }

            if (act == "type" && string.IsNullOrEmpty(step.Text))
            {
                return "type 必须提供 text";
            }

            if (act == "select" && string.IsNullOrWhiteSpace(step.Option))
            {
                return "select 必须提供 option";
            }

            return null;
        }

        private static BrowserInteractItemResult Succeeded(int index, BrowserInteractStep step, string refId)
        {
            return new BrowserInteractItemResult
            {
                ActionIndex = index,
                Action = (step.Action ?? "").Trim().ToLowerInvariant(),
                Ref = string.IsNullOrWhiteSpace(refId) ? step.Ref : refId,
                Status = "success",
            };
        }

        private static BrowserInteractItemResult Failed(int index, BrowserInteractStep step, string error)
        {
            return new BrowserInteractItemResult
            {
                ActionIndex = index,
                Action = (step.Action ?? "").Trim().ToLowerInvariant(),
                Ref = step.Ref,
                Status = "failed",
                Error = error ?? "interact 失败",
            };
        }

        private static BrowserInteractItemResult Skipped(int index, BrowserInteractStep step, string reason)
        {
            return new BrowserInteractItemResult
            {
                ActionIndex = index,
                Action = (step.Action ?? "").Trim().ToLowerInvariant(),
                Ref = step.Ref,
                Status = "skipped",
                Error = reason ?? "未执行",
            };
        }

        private static BrowserInteractStep ReadStep(Dictionary<string, object> row)
        {
            return new BrowserInteractStep
            {
                Action = ReadString(row, "action"),
                Ref = ReadString(row, "ref"),
                Text = ReadRaw(row, "text"),
                Direction = ReadString(row, "direction"),
                Key = ReadString(row, "key"),
                Option = ReadString(row, "option"),
            };
        }

        private static bool TryReadObjectList(
            object raw,
            out List<Dictionary<string, object>> rows,
            out string error)
        {
            rows = new List<Dictionary<string, object>>();
            error = null;
            if (raw is string)
            {
                error = "actions 须为对象数组";
                return false;
            }

            IEnumerable seq = raw as IEnumerable;
            if (seq == null)
            {
                error = "actions 须为对象数组";
                return false;
            }

            foreach (object item in seq)
            {
                Dictionary<string, object> row = CoerceRow(item);
                if (row == null)
                {
                    error = "actions 每一项须为对象";
                    return false;
                }

                rows.Add(row);
            }

            return true;
        }

        private static Dictionary<string, object> CoerceRow(object item)
        {
            if (item is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (item is IDictionary map)
            {
                var row = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
                foreach (DictionaryEntry entry in map)
                {
                    if (entry.Key != null)
                    {
                        row[Convert.ToString(entry.Key)] = entry.Value;
                    }
                }

                return row;
            }

            return null;
        }

        private static string ReadString(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return Convert.ToString(args[key])?.Trim();
        }

        private static string ReadRaw(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return null;
            }

            return Convert.ToString(args[key]);
        }
    }
}
