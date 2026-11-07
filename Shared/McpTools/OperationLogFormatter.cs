using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.Web.Script.Serialization;

namespace WordAddIn1
{
    public static class OperationLogFormatter
    {
        public const int MaxStringLength = 100;

        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer
        {
            RecursionLimit = 64
        };

        public static string FormatArgumentsJson(object arguments)
        {
            if (arguments == null)
            {
                return "{}";
            }

            object truncated = TruncateStrings(arguments);
            return Serializer.Serialize(truncated);
        }

        public static string FormatListLine(OperationLogEntry entry, int index, bool showBranch)
        {
            if (entry == null)
            {
                return $"{index}.";
            }

            if (entry.Kind == OperationLogEntryKind.Checkpoint)
            {
                if (entry.UserRollback)
                {
                    return null;
                }

                string line = $"{index}. 副本{entry.Timestamp}";
                if (entry.Branch > 0 && showBranch)
                {
                    line += " 分支" + entry.Branch;
                }

                return line;
            }

            string body = string.IsNullOrEmpty(entry.ActionLineBody) ? "{}" : entry.ActionLineBody;
            string toolName = entry.ToolName ?? "";
            return $"{index}. action: {toolName} {body}";
        }

        public static string FormatListText(DocumentCheckpointSession session)
        {
            if (session == null || session.Entries.Count == 0)
            {
                return "当前文档尚无操作记录。";
            }

            bool showBranch = session.Branches.Count > 1;
            var lines = new List<string>();
            int index = 1;
            for (int i = 0; i < session.Entries.Count; i++)
            {
                OperationLogEntry entry = session.Entries[i];
                if (entry != null
                    && entry.Kind == OperationLogEntryKind.Checkpoint
                    && entry.UserRollback)
                {
                    continue;
                }

                string line = FormatListLine(entry, index, showBranch);
                if (string.IsNullOrEmpty(line))
                {
                    continue;
                }

                lines.Add(line);
                index++;
            }

            if (lines.Count == 0)
            {
                return "当前文档尚无操作记录。";
            }

            return string.Join("\n", lines);
        }

        private static object TruncateStrings(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is string s)
            {
                return TruncateString(s);
            }

            if (value is IDictionary dict)
            {
                var result = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (DictionaryEntry pair in dict)
                {
                    string key = pair.Key?.ToString() ?? "";
                    result[key] = TruncateStrings(pair.Value);
                }

                return result;
            }

            if (value is IEnumerable list && !(value is string))
            {
                var items = new List<object>();
                foreach (object item in list)
                {
                    items.Add(TruncateStrings(item));
                }

                return items;
            }

            return value;
        }

        private static string TruncateString(string value)
        {
            if (string.IsNullOrEmpty(value) || value.Length <= MaxStringLength)
            {
                return value ?? "";
            }

            return value.Substring(0, MaxStringLength) + "...";
        }
    }
}
