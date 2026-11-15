using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>套用。先清这一侧的直接格式，再写样式。</summary>
    public static class StructureStyleApplier
    {
        public static Dictionary<string, object> Apply(
            Word.Document doc,
            Dictionary<string, object> item,
            IList<StructureStyleRecord> styles,
            bool disallowBackendApi)
        {
            string structure = Arg(item, "structure");
            if (string.IsNullOrWhiteSpace(structure))
            {
                return Fail("缺少 structure。");
            }

            StructureStyleRecord style = StructureStyleCatalog.Find(styles, structure);
            if (style == null)
            {
                return Fail("没有样式 " + structure + "。");
            }

            string code = Arg(item, "target_codes");
            if (string.IsNullOrWhiteSpace(code) || code.IndexOf(',') >= 0)
            {
                return Fail("target_codes 须是一个句子码。");
            }

            string sibling = TextSpanArg.RejectSiblingIndex(item);
            if (sibling != null)
            {
                return Fail(sibling);
            }

            var warnings = new List<string>();
            if (item != null && (item.ContainsKey("text_format") || item.ContainsKey("paragraph_format")))
            {
                warnings.Add("套用忽略 text_format 和 paragraph_format，按样式定义套。");
            }

            TextSpan span = null;
            if (item != null && item.ContainsKey("text"))
            {
                if (!TextSpanArg.TryRead(item, out span, out string parseError))
                {
                    return Fail(parseError ?? "text 无法解析。");
                }
            }

            string scope = Arg(item, "scope");
            if (string.IsNullOrEmpty(scope))
            {
                scope = span == null ? "paragraph" : "text";
            }

            if (style.Support == "paragraph")
            {
                scope = "paragraph";
            }

            InsertTargetLocatorHelper.ResolveRangeResult located = InsertTargetLocatorHelper.TryResolveCodeSpan(
                doc,
                new[] { code.Trim() },
                null,
                InsertTargetLocatorHelper.LocateRangeKind.ContentSpan,
                false,
                disallowBackendApi,
                span);
            if (located == null || !located.Success || located.Range == null)
            {
                return Fail(located == null ? "定位失败。" : located.Error);
            }

            try
            {
                Word.Range range = located.Range;
                Word.Paragraph paragraph = range.Paragraphs[1];
                if (style.Support == "both" && scope == "paragraph")
                {
                    ResetFont(ParagraphText(paragraph));
                    paragraph.Format.Reset();
                    SetStyle(paragraph.Range, Names(style, false));
                }
                else if (style.Support == "both" && style.Structure == "body")
                {
                    range.Font.Reset();
                }
                else if (style.Support == "both")
                {
                    range.Font.Reset();
                    SetStyle(range, Names(style, true));
                }
                else if (style.Support == "paragraph")
                {
                    paragraph.Format.Reset();
                    SetStyle(paragraph.Range, Names(style, false));
                }
                else
                {
                    Word.Range target = scope == "paragraph" ? ParagraphText(paragraph) : range;
                    target.Font.Reset();
                    SetStyle(target, Names(style, true));
                }
            }
            catch (Exception ex)
            {
                return Fail(ex.Message);
            }

            var done = new Dictionary<string, object>
            {
                ["ok"] = true,
                ["target_codes"] = code.Trim(),
                ["structure"] = style.Structure,
                ["scope"] = scope
            };
            if (warnings.Count > 0)
            {
                done["format_warnings"] = warnings;
            }

            return done;
        }

        private static Word.Range ParagraphText(Word.Paragraph paragraph)
        {
            Word.Range range = paragraph.Range.Duplicate;
            if (range.End > range.Start)
            {
                range.MoveEnd(Word.WdUnits.wdCharacter, -1);
            }

            return range;
        }

        private static void ResetFont(Word.Range range)
        {
            range.Font.Reset();
        }

        private static void SetStyle(Word.Range range, IList<string> names)
        {
            Exception last = null;
            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                try
                {
                    object styleRef = name;
                    range.set_Style(ref styleRef);
                    return;
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            throw last ?? new InvalidOperationException("样式名为空。");
        }

        private static IList<string> Names(StructureStyleRecord style, bool character)
        {
            var names = new List<string>();
            if (character)
            {
                Add(names, style.CharacterStyleName);
                Add(names, style.StyleName);
            }
            else
            {
                Add(names, style.StyleName);
            }

            foreach (string name in StructureStyleCatalog.LookupNames(style))
            {
                Add(names, name);
            }

            return names;
        }

        private static void Add(List<string> names, string name)
        {
            if (!string.IsNullOrWhiteSpace(name) && !names.Contains(name))
            {
                names.Add(name);
            }
        }

        private static string Arg(Dictionary<string, object> item, string key)
        {
            if (item == null || !item.ContainsKey(key) || item[key] == null)
            {
                return null;
            }

            return item[key].ToString().Trim();
        }

        private static Dictionary<string, object> Fail(string error)
        {
            return new Dictionary<string, object> { ["ok"] = false, ["error"] = error };
        }
    }
}
