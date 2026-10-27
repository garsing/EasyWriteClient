using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 把一次套用里的多段格式，按当前 seg 快照拆成统一定位能接受的连续子操作。
    /// </summary>
    public static class FormatApplyBatchCompiler
    {
        public sealed class CompiledOp
        {
            public int ItemIndex { get; set; }
            public bool Paragraph { get; set; }
            public List<string> Codes { get; set; }
            public Dictionary<string, object> Format { get; set; }
        }

        public sealed class CompileResult
        {
            public string ContractError { get; set; }
            public List<CompiledOp> Ops { get; set; } = new List<CompiledOp>();
            public List<Dictionary<string, object>> Warnings { get; set; } = new List<Dictionary<string, object>>();
        }

        public static bool HasItems(Dictionary<string, object> args)
        {
            return args != null && args.ContainsKey("items");
        }

        public static string ValidateDocumentItems(Dictionary<string, object> args)
        {
            if (ActionFormatHelper.HasCancelledInheritKeys(args))
            {
                return ActionFormatHelper.CancelledInheritFromError;
            }

            if (!TryReadItems(args, out List<Dictionary<string, object>> items, out string error))
            {
                return error;
            }

            for (int i = 0; i < items.Count; i++)
            {
                Dictionary<string, object> item = items[i];
                if (item == null)
                {
                    return "items[" + i + "] 不是对象";
                }

                if (HasText(item, "target_codes") && HasText(item, "target_paragraph_codes"))
                {
                    return "items[" + i + "] 的 target_codes 与 target_paragraph_codes 互斥";
                }

                if (!TryReadCodes(item, out List<string> codes, out bool paragraph, out string codeError))
                {
                    return "items[" + i + "] " + codeError;
                }

                if (codes.Count == 0)
                {
                    return "items[" + i + "] 缺少编码";
                }

                if (codes.Any(IsTableCode))
                {
                    return "items[" + i + "] 仅支持 S_ / P_，不支持 T_";
                }

                Dictionary<string, object> format = ReadObject(item, "char_format");
                if (format == null || format.Count == 0)
                {
                    return "items[" + i + "] 缺少 char_format";
                }

                bool anyP = codes.Any(IsParagraphCode);
                bool anyS = codes.Any(IsSentenceCode);
                if (anyP && anyS)
                {
                    continue;
                }

                if (!anyP && !anyS)
                {
                    return "items[" + i + "] 仅支持 S_ / P_";
                }

                if (paragraph && anyS)
                {
                    return "items[" + i + "] 的 target_paragraph_codes 只能是 P_";
                }
            }

            return null;
        }

        public static string ValidateParagraphItems(Dictionary<string, object> args)
        {
            if (ActionFormatHelper.HasCancelledInheritKeys(args))
            {
                return ActionFormatHelper.CancelledInheritFromParagraphError;
            }

            if (!TryReadItems(args, out List<Dictionary<string, object>> items, out string error))
            {
                return error;
            }

            for (int i = 0; i < items.Count; i++)
            {
                Dictionary<string, object> item = items[i];
                if (item == null)
                {
                    return "items[" + i + "] 不是对象";
                }

                if (HasText(item, "target_codes"))
                {
                    return "items[" + i + "] 请使用 target_paragraph_codes";
                }

                List<string> codes = SplitCodes(Text(item, "target_paragraph_codes"));
                if (codes.Count == 0)
                {
                    return "items[" + i + "] 缺少 target_paragraph_codes";
                }

                if (codes.Any(code => !IsParagraphCode(code)))
                {
                    return "items[" + i + "] 只接受 P_";
                }

                Dictionary<string, object> format = ReadObject(item, "para_format");
                if (format == null || format.Count == 0)
                {
                    return "items[" + i + "] 缺少 para_format";
                }
            }

            return null;
        }

        public static CompileResult CompileDocument(Dictionary<string, object> args)
        {
            string contract = HasItems(args) ? ValidateDocumentItems(args) : null;
            if (contract != null)
            {
                return new CompileResult { ContractError = contract };
            }

            if (!TryReadDocumentWorkItems(args, out List<Dictionary<string, object>> items, out string error))
            {
                return new CompileResult { ContractError = error };
            }

            return Compile(items, "char_format", paragraphTool: false);
        }

        public static CompileResult CompileParagraph(Dictionary<string, object> args)
        {
            string contract = HasItems(args) ? ValidateParagraphItems(args) : null;
            if (contract != null)
            {
                return new CompileResult { ContractError = contract };
            }

            if (!TryReadParagraphWorkItems(args, out List<Dictionary<string, object>> items, out string error))
            {
                return new CompileResult { ContractError = error };
            }

            return Compile(items, "para_format", paragraphTool: true);
        }

        public static List<List<string>> SplitRuns(IList<string> codes, string segSnapshot)
        {
            var runs = new List<List<string>>();
            if (codes == null || codes.Count == 0)
            {
                return runs;
            }

            List<string[]> segs = ParseSegs(segSnapshot);
            int offset = 0;
            while (offset < codes.Count)
            {
                int best = 1;
                List<Pos> hits = Find(segs, codes[offset]);
                if (hits.Count > 0)
                {
                    best = 0;
                    foreach (Pos hit in hits)
                    {
                        int len = MatchLength(segs, hit, codes, offset);
                        if (len > best)
                        {
                            best = len;
                        }
                    }

                    if (best < 1)
                    {
                        best = 1;
                    }
                }

                runs.Add(codes.Skip(offset).Take(best).ToList());
                offset += best;
            }

            return runs;
        }

        private static CompileResult Compile(
            List<Dictionary<string, object>> items,
            string formatKey,
            bool paragraphTool)
        {
            var result = new CompileResult();
            for (int i = 0; i < items.Count; i++)
            {
                Dictionary<string, object> item = items[i];
                Dictionary<string, object> format = ReadObject(item, formatKey);
                if (!TryReadCodes(item, out List<string> codes, out bool paragraph, out _))
                {
                    result.Warnings.Add(Warning(i, "", "缺少编码，未套用"));
                    continue;
                }

                bool anyP = codes.Any(IsParagraphCode);
                bool anyS = codes.Any(IsSentenceCode);
                if (anyP && anyS)
                {
                    result.Warnings.Add(Warning(i, string.Join(",", codes), "同一项混用 S_ 与 P_，未套用"));
                    continue;
                }

                if (paragraphTool)
                {
                    paragraph = true;
                }

                string segSnapshot = paragraph || anyP
                    ? DocumentState.ParagraphSegSnapshot
                    : DocumentState.SegSnapshot;
                foreach (List<string> run in SplitRuns(codes, segSnapshot))
                {
                    result.Ops.Add(new CompiledOp
                    {
                        ItemIndex = i,
                        Paragraph = paragraph || anyP,
                        Codes = run,
                        Format = format
                    });
                }
            }

            return result;
        }

        private static bool TryReadDocumentWorkItems(
            Dictionary<string, object> args,
            out List<Dictionary<string, object>> items,
            out string error)
        {
            if (HasItems(args))
            {
                return TryReadItems(args, out items, out error);
            }

            items = new List<Dictionary<string, object>>();
            error = null;
            var flat = new Dictionary<string, object>();
            CopyIfPresent(args, flat, "target_codes");
            CopyIfPresent(args, flat, "target_paragraph_codes");
            CopyIfPresent(args, flat, "char_format");
            items.Add(flat);
            return true;
        }

        private static bool TryReadParagraphWorkItems(
            Dictionary<string, object> args,
            out List<Dictionary<string, object>> items,
            out string error)
        {
            if (HasItems(args))
            {
                return TryReadItems(args, out items, out error);
            }

            items = new List<Dictionary<string, object>>();
            error = null;
            if (!HasText(args, "target_paragraph_codes"))
            {
                return true;
            }

            var flat = new Dictionary<string, object>();
            CopyIfPresent(args, flat, "target_paragraph_codes");
            CopyIfPresent(args, flat, "para_format");
            items.Add(flat);
            return true;
        }

        private static bool TryReadItems(
            Dictionary<string, object> args,
            out List<Dictionary<string, object>> items,
            out string error)
        {
            items = new List<Dictionary<string, object>>();
            error = null;
            if (args == null || !args.TryGetValue("items", out object raw) || raw == null)
            {
                error = "items 为空";
                return false;
            }

            foreach (object entry in Enumerate(raw))
            {
                items.Add(AsDictionary(entry));
            }

            if (items.Count == 0)
            {
                error = "items 为空";
                return false;
            }

            return true;
        }

        private static bool TryReadCodes(
            Dictionary<string, object> item,
            out List<string> codes,
            out bool paragraph,
            out string error)
        {
            codes = new List<string>();
            paragraph = false;
            error = null;
            if (item == null)
            {
                error = "不是对象";
                return false;
            }

            if (HasText(item, "target_paragraph_codes"))
            {
                paragraph = true;
                codes = SplitCodes(Text(item, "target_paragraph_codes"));
                return true;
            }

            if (HasText(item, "target_codes"))
            {
                codes = SplitCodes(Text(item, "target_codes"));
                return true;
            }

            error = "缺少 target_codes 或 target_paragraph_codes";
            return false;
        }

        public static Dictionary<string, object> Warning(int itemIndex, string codes, string message)
        {
            return new Dictionary<string, object>
            {
                ["item_index"] = itemIndex,
                ["codes"] = codes ?? "",
                ["message"] = message ?? ""
            };
        }

        public static Dictionary<string, object> CopyArgsForOp(
            Dictionary<string, object> args,
            CompiledOp op,
            string formatKey)
        {
            var one = new Dictionary<string, object>();
            if (args != null)
            {
                foreach (KeyValuePair<string, object> pair in args)
                {
                    if (pair.Key == "items"
                        || pair.Key == "target_codes"
                        || pair.Key == "target_paragraph_codes"
                        || pair.Key == formatKey)
                    {
                        continue;
                    }

                    one[pair.Key] = pair.Value;
                }
            }

            string csv = string.Join(",", op.Codes ?? new List<string>());
            if (op.Paragraph)
            {
                one["target_paragraph_codes"] = csv;
            }
            else
            {
                one["target_codes"] = csv;
            }

            if (op.Format != null)
            {
                one[formatKey] = op.Format;
            }

            return one;
        }

        private static List<string[]> ParseSegs(string segSnapshot)
        {
            var segs = new List<string[]>();
            if (string.IsNullOrEmpty(segSnapshot))
            {
                return segs;
            }

            foreach (string part in segSnapshot.Split(';'))
            {
                segs.Add(SplitCodes(part).ToArray());
            }

            return segs;
        }

        private static List<Pos> Find(List<string[]> segs, string code)
        {
            var hits = new List<Pos>();
            if (string.IsNullOrEmpty(code))
            {
                return hits;
            }

            for (int seg = 0; seg < segs.Count; seg++)
            {
                string[] codes = segs[seg];
                for (int index = 0; index < codes.Length; index++)
                {
                    if (string.Equals(codes[index], code, StringComparison.Ordinal))
                    {
                        hits.Add(new Pos { Seg = seg, Index = index });
                    }
                }
            }

            return hits;
        }

        private static int MatchLength(List<string[]> segs, Pos start, IList<string> codes, int offset)
        {
            string[] seg = segs[start.Seg];
            int matched = 0;
            while (offset + matched < codes.Count
                && start.Index + matched < seg.Length
                && string.Equals(seg[start.Index + matched], codes[offset + matched], StringComparison.Ordinal))
            {
                matched++;
            }

            return matched;
        }

        private static List<string> SplitCodes(string csv)
        {
            var codes = new List<string>();
            if (string.IsNullOrWhiteSpace(csv))
            {
                return codes;
            }

            foreach (string part in csv.Split(','))
            {
                string code = part.Trim();
                if (code.Length > 0)
                {
                    codes.Add(code);
                }
            }

            return codes;
        }

        private static IEnumerable Enumerate(object raw)
        {
            if (raw is string)
            {
                yield break;
            }

            if (raw is IEnumerable enumerable)
            {
                foreach (object entry in enumerable)
                {
                    yield return entry;
                }
            }
        }

        private static Dictionary<string, object> AsDictionary(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (raw is IDictionary generic)
            {
                var copy = new Dictionary<string, object>();
                foreach (DictionaryEntry entry in generic)
                {
                    copy[entry.Key?.ToString() ?? ""] = entry.Value;
                }

                return copy;
            }

            return null;
        }

        private static Dictionary<string, object> ReadObject(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.TryGetValue(key, out object raw) || raw == null)
            {
                return null;
            }

            return AsDictionary(raw);
        }

        private static bool HasText(Dictionary<string, object> args, string key)
        {
            return !string.IsNullOrWhiteSpace(Text(args, key));
        }

        private static string Text(Dictionary<string, object> args, string key)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return "";
            }

            return args[key].ToString().Trim();
        }

        private static void CopyIfPresent(
            Dictionary<string, object> source,
            Dictionary<string, object> target,
            string key)
        {
            if (source != null && source.ContainsKey(key))
            {
                target[key] = source[key];
            }
        }

        private static bool IsParagraphCode(string code)
        {
            return code != null && code.StartsWith("P_", StringComparison.Ordinal);
        }

        private static bool IsSentenceCode(string code)
        {
            return code != null && code.StartsWith("S_", StringComparison.Ordinal);
        }

        private static bool IsTableCode(string code)
        {
            return code != null
                && code.StartsWith("T_", StringComparison.Ordinal)
                && !EmptyCellCode.IsEmptyCellCode(code);
        }

        private struct Pos
        {
            public int Seg;
            public int Index;
        }
    }
}
