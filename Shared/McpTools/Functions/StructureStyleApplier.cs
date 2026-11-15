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

            if (style.Support == "paragraph" && scope == "text")
            {
                return Done(code, style, scope, warnings, ParagraphSkippedNote);
            }

            try
            {
                Word.Range range = located.Range;
                Word.Paragraph paragraph = range.Paragraphs[1];
                if (style.Support == "both" && scope == "paragraph")
                {
                    ResetFont(ParagraphText(paragraph));
                    paragraph.Format.Reset();
                    SetStyle(doc, paragraph.Range, Names(style, false), false);
                }
                else if (style.Support == "both" && style.Structure == "body")
                {
                    range.Font.Reset();
                }
                else if (style.Support == "both")
                {
                    range.Font.Reset();
                    SetStyle(doc, range, Names(style, true), true);
                }
                else if (style.Support == "paragraph")
                {
                    paragraph.Format.Reset();
                    SetStyle(doc, paragraph.Range, Names(style, false), false);
                }
                else
                {
                    Word.Range target = scope == "paragraph" ? ParagraphText(paragraph) : range;
                    target.Font.Reset();
                    SetStyle(doc, target, Names(style, true), true);
                }
            }
            catch (Exception ex)
            {
                return Fail(ex.Message);
            }

            string note = scope == "text" && style.Support == "both" ? CharacterApplyNote : null;
            return Done(code, style, scope, warnings, note);
        }

        private const string CharacterApplyNote = "由于是字符套用，所以段落格式没有发生作用。";

        private const string ParagraphSkippedNote =
            "由于是字符套用，所以段落格式没有发生作用。字符范围套到了只有段落格式的样式上，没有字符格式可套，所以这次套用什么也不会发生。";

        private static Dictionary<string, object> Done(
            string code,
            StructureStyleRecord style,
            string scope,
            List<string> warnings,
            string note)
        {
            var done = new Dictionary<string, object>
            {
                ["ok"] = true,
                ["target_codes"] = code.Trim(),
                ["structure"] = style.Structure,
                ["scope"] = scope
            };
            if (!string.IsNullOrEmpty(note))
            {
                done["note"] = note;
            }

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

        private static void SetStyle(Word.Document doc, Word.Range range, IList<string> names, bool character)
        {
            Exception last = null;
            foreach (string name in names)
            {
                if (string.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                Word.Style found = FindStyle(doc, name);
                if (found == null)
                {
                    continue;
                }

                bool isCharacter = found.Type == Word.WdStyleType.wdStyleTypeCharacter;
                if (isCharacter != character)
                {
                    continue;
                }

                try
                {
                    object styleRef = found;
                    range.set_Style(ref styleRef);
                    return;
                }
                catch (Exception ex)
                {
                    last = ex;
                }
            }

            throw last ?? new InvalidOperationException("文档里找不到对应的样式。");
        }

        private static Word.Style FindStyle(Word.Document doc, string name)
        {
            try
            {
                return doc.Styles[name];
            }
            catch (Exception)
            {
                return null;
            }
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
