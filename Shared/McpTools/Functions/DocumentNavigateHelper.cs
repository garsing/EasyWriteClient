using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class NavigateActionInput
    {
        public int ActionIndex { get; set; }
        public string Type { get; set; }
        public bool Success { get; set; }
        public List<string> AffectedCodes { get; set; }
        public List<string> NewCodes { get; set; }
        public List<Dictionary<string, object>> NewSentences { get; set; }
    }

    public static class DocumentNavigateHelper
    {
        public const string SelectionModeScrollToCursor = "scroll_to_cursor";

        /// <summary>
        /// insert/replace/delete 执行时已移动光标；此处仅将视口滚到当前选区，不再按编码二次查找。
        /// </summary>
        public static (List<Dictionary<string, object>> Navigates, Dictionary<string, object> Summary) BuildAndApply(
            Word.Application app,
            Word.Document doc,
            IReadOnlyList<NavigateActionInput> actions,
            bool bulkMode,
            bool autoNavigateEnabled)
        {
            var navigates = new List<Dictionary<string, object>>();
            var navigable = (actions ?? Array.Empty<NavigateActionInput>())
                .Where(a => a != null && a.Success && IsNavigableType(a.Type))
                .OrderBy(a => a.ActionIndex)
                .ToList();

            int? lastNavigableIndex = navigable.Count > 0
                ? navigable[navigable.Count - 1].ActionIndex
                : (int?)null;

            foreach (NavigateActionInput action in navigable)
            {
                bool isLast = lastNavigableIndex.HasValue && action.ActionIndex == lastNavigableIndex.Value;
                Dictionary<string, object> entry = BuildBaseEntry(action);
                entry["anchor"] = "current_selection";

                if (!autoNavigateEnabled)
                {
                    MarkSkipped(entry, "disabled_by_option");
                    navigates.Add(entry);
                    continue;
                }

                if (bulkMode)
                {
                    MarkSkipped(entry, "bulk_snapshot");
                    navigates.Add(entry);
                    continue;
                }

                if (!isLast)
                {
                    entry["status"] = "computed";
                    entry["applied"] = false;
                    entry["reason"] = "not_last_action";
                    navigates.Add(entry);
                    continue;
                }

                ApplyScrollToCurrentSelection(app, entry);
                navigates.Add(entry);
            }

            return (navigates, BuildSummary(navigates, lastNavigableIndex));
        }

        private static bool IsNavigableType(string type)
        {
            if (string.IsNullOrEmpty(type))
            {
                return false;
            }

            string normalized = type.ToLowerInvariant();
            return normalized == "replace" || normalized == "insert" || normalized == "delete";
        }

        private static Dictionary<string, object> BuildBaseEntry(NavigateActionInput action)
        {
            return new Dictionary<string, object>
            {
                ["action_index"] = action.ActionIndex,
                ["type"] = action.Type,
                ["status"] = "skipped",
                ["applied"] = false,
                ["selection_mode"] = SelectionModeScrollToCursor
            };
        }

        private static void MarkSkipped(Dictionary<string, object> entry, string reason)
        {
            entry["status"] = "skipped";
            entry["applied"] = false;
            entry["reason"] = reason;
        }

        private static void ApplyScrollToCurrentSelection(Word.Application app, Dictionary<string, object> entry)
        {
            if (app?.Selection == null)
            {
                entry["status"] = "failed";
                entry["applied"] = false;
                entry["reason"] = "no_selection";
                return;
            }

            try
            {
                Word.Range range = app.Selection.Range;
                if (range == null)
                {
                    entry["status"] = "failed";
                    entry["applied"] = false;
                    entry["reason"] = "no_selection";
                    return;
                }

                entry["range"] = new Dictionary<string, object>
                {
                    ["start"] = range.Start,
                    ["end"] = range.End
                };

                if (!WordRangeFinder.ScrollRangeIntoView(app, range, "current_selection"))
                {
                    entry["status"] = "failed";
                    entry["applied"] = false;
                    entry["reason"] = "scroll_failed";
                    return;
                }

                entry["status"] = "ok";
                entry["applied"] = true;
                entry["reason"] = null;
            }
            catch (Exception ex)
            {
                entry["status"] = "failed";
                entry["applied"] = false;
                entry["reason"] = "scroll_failed";
                System.Diagnostics.Debug.WriteLine($"[DocumentNavigate] 滚动到当前光标失败: {ex.Message}");
            }
        }

        private static Dictionary<string, object> BuildSummary(
            List<Dictionary<string, object>> navigates,
            int? lastNavigableIndex)
        {
            Dictionary<string, object> applied = navigates?.FirstOrDefault(e =>
                e.TryGetValue("applied", out object appliedObj)
                && appliedObj is bool appliedBool
                && appliedBool);

            if (applied != null)
            {
                return CloneSummary(applied);
            }

            Dictionary<string, object> last = navigates?.LastOrDefault(e =>
                lastNavigableIndex.HasValue
                && e.TryGetValue("action_index", out object idxObj)
                && Convert.ToInt32(idxObj) == lastNavigableIndex.Value);

            if (last != null)
            {
                return CloneSummary(last);
            }

            return new Dictionary<string, object>
            {
                ["status"] = "skipped",
                ["applied"] = false,
                ["selection_mode"] = SelectionModeScrollToCursor,
                ["reason"] = "no_navigable_action"
            };
        }

        private static Dictionary<string, object> CloneSummary(Dictionary<string, object> source)
        {
            var summary = new Dictionary<string, object>();
            foreach (var kvp in source)
            {
                summary[kvp.Key] = kvp.Value;
            }

            return summary;
        }
    }
}
