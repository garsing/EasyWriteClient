using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// I13：COM 读各节 PageSetup + 页眉脚 Story，输出对齐 KB section_page_layout。
    /// </summary>
    public static class PageLayoutExtractor
    {
        public static Dictionary<string, object> Extract(Word.Document document)
        {
            var warnings = new List<string>();
            var sections = new List<Dictionary<string, object>>();
            if (document == null)
            {
                return new Dictionary<string, object>
                {
                    ["version"] = 1,
                    ["section_count"] = 0,
                    ["sections"] = sections,
                    ["extract_warnings"] = new List<string> { "document 为空" }
                };
            }

            int count;
            try
            {
                count = document.Sections.Count;
            }
            catch (Exception ex)
            {
                return new Dictionary<string, object>
                {
                    ["version"] = 1,
                    ["section_count"] = 0,
                    ["sections"] = sections,
                    ["extract_warnings"] = new List<string> { "无法读取节: " + ex.Message }
                };
            }

            for (int i = 1; i <= count; i++)
            {
                try
                {
                    Word.Section section = document.Sections[i];
                    sections.Add(ExtractSection(section, i));
                }
                catch (Exception ex)
                {
                    warnings.Add("section " + i + " extract failed: " + ex.Message);
                    sections.Add(EmptySection(i));
                }
            }

            var payload = new Dictionary<string, object>
            {
                ["version"] = 1,
                ["section_count"] = sections.Count,
                ["sections"] = sections
            };

            if (sections.Count > 0 && sections[0].TryGetValue("page_setup", out object firstPs))
            {
                var mode = new Dictionary<string, object>
                {
                    ["page_setup"] = CloneDict(firstPs as Dictionary<string, object>),
                    ["headers"] = CloneDict(GetDict(sections[0], "headers")),
                    ["footers"] = CloneDict(GetDict(sections[0], "footers")),
                    ["mode_source_section_index"] = 1,
                    ["mode_occurrence_count"] = sections.Count
                };
                payload["mode"] = mode;
            }

            if (warnings.Count > 0)
            {
                payload["extract_warnings"] = warnings;
            }

            return payload;
        }

        private static Dictionary<string, object> ExtractSection(Word.Section section, int index)
        {
            Word.PageSetup ps = section.PageSetup;
            var pageSetup = new Dictionary<string, object>
            {
                ["orientation"] = ps.Orientation == Word.WdOrientation.wdOrientLandscape
                    ? "landscape"
                    : "portrait",
                ["page_width"] = RoundPt(ps.PageWidth),
                ["page_height"] = RoundPt(ps.PageHeight),
                ["top_margin"] = RoundPt(ps.TopMargin),
                ["bottom_margin"] = RoundPt(ps.BottomMargin),
                ["left_margin"] = RoundPt(ps.LeftMargin),
                ["right_margin"] = RoundPt(ps.RightMargin),
                ["header_distance"] = RoundPt(ps.HeaderDistance),
                ["footer_distance"] = RoundPt(ps.FooterDistance),
                ["different_first_page_header_footer"] = ToBool(ps.DifferentFirstPageHeaderFooter),
                ["odd_and_even_pages_header_footer"] = ToBool(ps.OddAndEvenPagesHeaderFooter)
            };

            return new Dictionary<string, object>
            {
                ["section_index"] = index,
                ["pages_in_section"] = 1,
                ["extract_origin"] = "com",
                ["page_setup"] = pageSetup,
                ["headers"] = ExtractStories(section, header: true),
                ["footers"] = ExtractStories(section, header: false)
            };
        }

        private static Dictionary<string, object> ExtractStories(Word.Section section, bool header)
        {
            var stories = new Dictionary<string, object>(StringComparer.Ordinal);
            TryAddStory(stories, "primary", GetStory(section, header, Word.WdHeaderFooterIndex.wdHeaderFooterPrimary));
            TryAddStory(stories, "first", GetStory(section, header, Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage));
            TryAddStory(stories, "even", GetStory(section, header, Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages));
            return stories.Count == 0 ? null : stories;
        }

        private static void TryAddStory(
            Dictionary<string, object> stories,
            string key,
            Dictionary<string, object> story)
        {
            if (story != null && story.Count > 0)
            {
                stories[key] = story;
            }
        }

        private static Dictionary<string, object> GetStory(
            Word.Section section,
            bool header,
            Word.WdHeaderFooterIndex index)
        {
            try
            {
                Word.HeaderFooter hf = header ? section.Headers[index] : section.Footers[index];
                if (hf == null || hf.LinkToPrevious)
                {
                    return null;
                }

                Word.Range range = hf.Range;
                string text = (range.Text ?? "").Replace("\r", "").Trim();
                var story = new Dictionary<string, object>();
                if (!string.IsNullOrEmpty(text))
                {
                    story["text"] = text;
                }

                try
                {
                    Word.Font font = range.Font;
                    var charFormat = new Dictionary<string, object>();
                    if (!string.IsNullOrEmpty(font.Name))
                    {
                        charFormat["font_name"] = font.Name;
                    }

                    if (font.Size > 0)
                    {
                        charFormat["font_size"] = (double)font.Size;
                    }

                    if (charFormat.Count > 0)
                    {
                        story["char_format"] = charFormat;
                    }
                }
                catch
                {
                }

                return story.Count == 0 ? null : story;
            }
            catch
            {
                return null;
            }
        }

        private static Dictionary<string, object> EmptySection(int index)
        {
            return new Dictionary<string, object>
            {
                ["section_index"] = index,
                ["pages_in_section"] = 1,
                ["extract_origin"] = "empty",
                ["page_setup"] = new Dictionary<string, object>(),
                ["headers"] = null,
                ["footers"] = null
            };
        }

        private static double RoundPt(float value)
        {
            return Math.Round(value, 2);
        }

        private static bool ToBool(int wordInt)
        {
            return wordInt != 0;
        }

        private static Dictionary<string, object> GetDict(Dictionary<string, object> row, string key)
        {
            if (row != null && row.TryGetValue(key, out object raw) && raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            return null;
        }

        private static Dictionary<string, object> CloneDict(Dictionary<string, object> src)
        {
            if (src == null)
            {
                return null;
            }

            return new Dictionary<string, object>(src, StringComparer.Ordinal);
        }
    }
}
