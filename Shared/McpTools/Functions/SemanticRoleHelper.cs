using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    /// <summary>
    /// Phase 2 PR-1b：semantic_role 双源合成（Word Style → style_role + HeadingRecognizer → text_role，方案 C）。
    /// </summary>
    public static class SemanticRoleHelper
    {
        private static readonly HeadingRecognizer Recognizer = new HeadingRecognizer();

        private static readonly Dictionary<string, string> StyleRoleMap =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["正文"] = "body",
                ["Normal"] = "body",
                ["正文 Char"] = "body",
                ["标题 1"] = "heading_l1",
                ["标题 2"] = "heading_l2",
                ["标题 3"] = "heading_l3",
                ["标题 4"] = "heading_l4",
                ["标题 5"] = "heading_l5",
                ["标题 6"] = "heading_l6",
                ["标题 7"] = "heading_l7",
                ["标题 8"] = "heading_l8",
                ["标题 9"] = "heading_l9",
                ["Heading 1"] = "heading_l1",
                ["Heading 2"] = "heading_l2",
                ["Heading 3"] = "heading_l3",
                ["Heading 4"] = "heading_l4",
                ["Heading 5"] = "heading_l5",
                ["Heading 6"] = "heading_l6",
                ["Heading 7"] = "heading_l7",
                ["Heading 8"] = "heading_l8",
                ["Heading 9"] = "heading_l9",
                ["题注"] = "caption",
                ["Caption"] = "caption",
                ["列表段落"] = "list_item",
                ["List Paragraph"] = "list_item",
            };

        public sealed class RoleAnalysis
        {
            public string StyleRole { get; set; } = "body";
            public string TextRole { get; set; } = "body";
            public string SemanticRole { get; set; }
            public bool RoleConflict { get; set; }
            public string RoleSource { get; set; }
            public string Confidence { get; set; } = "high";
            public string StyleRoleNote { get; set; }
        }

        public static RoleAnalysis Analyze(string sentenceCode, string content, Dictionary<string, object> snapshot)
        {
            string paragraphStyle = snapshot != null && snapshot.TryGetValue("paragraph_style", out object ps)
                ? ps?.ToString() ?? ""
                : "";

            var styleSide = MapParagraphStyleToStyleRole(paragraphStyle);
            var textSide = InferTextRole(sentenceCode, content);
            return Synthesize(styleSide, textSide);
        }

        public static void AppendRoleFields(Dictionary<string, object> entry, RoleAnalysis role)
        {
            if (entry == null || role == null)
            {
                return;
            }

            entry["style_role"] = role.StyleRole ?? "body";
            entry["text_role"] = role.TextRole ?? "body";
            entry["role_conflict"] = role.RoleConflict;
            entry["confidence"] = role.Confidence ?? "high";

            if (!string.IsNullOrEmpty(role.SemanticRole))
            {
                entry["semantic_role"] = role.SemanticRole;
            }

            if (!string.IsNullOrEmpty(role.RoleSource))
            {
                entry["role_source"] = role.RoleSource;
            }

            if (!string.IsNullOrEmpty(role.StyleRoleNote))
            {
                entry["style_role_note"] = role.StyleRoleNote;
            }
        }

        private static Tuple<string, string, bool> MapParagraphStyleToStyleRole(string paragraphStyle)
        {
            string styleName = (paragraphStyle ?? "").Trim();
            if (string.IsNullOrEmpty(styleName))
            {
                return Tuple.Create("body", (string)null, false);
            }

            if (StyleRoleMap.TryGetValue(styleName, out string mapped))
            {
                return Tuple.Create(mapped, (string)null, true);
            }

            Match headingMatch = Regex.Match(
                styleName,
                @"^(?:标题\s*(\d)|Heading\s*(\d))$",
                RegexOptions.IgnoreCase);
            if (headingMatch.Success)
            {
                string levelStr = !string.IsNullOrEmpty(headingMatch.Groups[1].Value)
                    ? headingMatch.Groups[1].Value
                    : headingMatch.Groups[2].Value;
                if (int.TryParse(levelStr, out int level))
                {
                    return Tuple.Create(MapHeadingLevelToRole(level), (string)null, true);
                }
            }

            if (styleName.IndexOf("caption", StringComparison.OrdinalIgnoreCase) >= 0
                || styleName.IndexOf("题注", StringComparison.Ordinal) >= 0)
            {
                return Tuple.Create("caption", (string)null, true);
            }

            if (styleName.IndexOf("list", StringComparison.OrdinalIgnoreCase) >= 0
                || styleName.IndexOf("列表", StringComparison.Ordinal) >= 0)
            {
                return Tuple.Create("list_item", (string)null, true);
            }

            return Tuple.Create("body", $"unmapped_style:{styleName}", false);
        }

        private static Tuple<string, string> InferTextRole(string sentenceCode, string content)
        {
            if (TryGetTextRoleFromHeadingMap(sentenceCode, out string mappedRole, out string note))
            {
                return Tuple.Create(mappedRole, note);
            }

            if (string.IsNullOrWhiteSpace(content))
            {
                return Tuple.Create("body", "text_empty");
            }

            Tuple<bool, string, string> heading = Recognizer.IsHeading(content);
            if (!heading.Item1)
            {
                return Tuple.Create("body", "text_role_confidence:low");
            }

            if (string.Equals(heading.Item2, "markdown", StringComparison.OrdinalIgnoreCase)
                || string.Equals(heading.Item3, "markdown", StringComparison.OrdinalIgnoreCase))
            {
                int mdLevel = GetMarkdownHeadingLevel(content);
                return Tuple.Create(MapHeadingLevelToRole(mdLevel), (string)null);
            }

            string langType = heading.Item2 ?? "";
            string headingType = heading.Item3 ?? "";
            string trimmed = content.TrimEnd().Trim();
            string detailedSubtype = Recognizer.GetDetailedSubtype(trimmed, langType, headingType);
            int level = InferLevelFromDetailedSubtype(detailedSubtype, headingType);
            return Tuple.Create(MapHeadingLevelToRole(level), (string)null);
        }

        private static bool TryGetTextRoleFromHeadingMap(string sentenceCode, out string textRole, out string note)
        {
            textRole = "body";
            note = null;

            if (string.IsNullOrEmpty(sentenceCode))
            {
                return false;
            }

            Dictionary<string, Dictionary<string, object>> map = DocumentState.SentenceNameToHeadingMap;
            if (map == null || !map.TryGetValue(sentenceCode, out Dictionary<string, object> info) || info == null)
            {
                return false;
            }

            string type = info.TryGetValue("type", out object t) ? t?.ToString() ?? "" : "";
            if (!string.Equals(type, "heading", StringComparison.OrdinalIgnoreCase))
            {
                textRole = "body";
                return true;
            }

            if (info.TryGetValue("level", out object levelObj))
            {
                int level = ParseHeadingLevelToken(levelObj?.ToString());
                if (level > 0)
                {
                    textRole = MapHeadingLevelToRole(level);
                    return true;
                }
            }

            if (info.TryGetValue("detailed_subtype", out object dstObj))
            {
                string dst = dstObj?.ToString() ?? "";
                textRole = MapHeadingLevelToRole(InferLevelFromDetailedSubtype(dst, ""));
                return true;
            }

            textRole = "heading_l1";
            return true;
        }

        private static RoleAnalysis Synthesize(
            Tuple<string, string, bool> styleSide,
            Tuple<string, string> textSide)
        {
            string styleRole = styleSide.Item1 ?? "body";
            string textRole = textSide.Item1 ?? "body";
            bool styleMapped = styleSide.Item3;
            string styleNote = styleSide.Item2;

            var analysis = new RoleAnalysis
            {
                StyleRole = styleRole,
                TextRole = textRole,
                StyleRoleNote = styleNote
            };

            if (string.Equals(styleRole, textRole, StringComparison.Ordinal))
            {
                analysis.SemanticRole = styleRole;
                analysis.RoleConflict = false;
                analysis.RoleSource = "both";
                analysis.Confidence = (!styleMapped || !string.IsNullOrEmpty(styleNote)) ? "medium" : "high";
                return analysis;
            }

            analysis.SemanticRole = null;
            analysis.RoleConflict = true;
            analysis.RoleSource = "conflict";
            analysis.Confidence = "low";
            return analysis;
        }

        private static int InferLevelFromDetailedSubtype(string detailedSubtype, string headingType)
        {
            string dst = detailedSubtype ?? "";
            if (dst.Contains("level3"))
            {
                return 3;
            }

            if (dst.Contains("level2"))
            {
                return 2;
            }

            if (string.Equals(headingType, "chapter", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            return 1;
        }

        private static int GetMarkdownHeadingLevel(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return 1;
            }

            int count = 0;
            foreach (char c in text.TrimStart())
            {
                if (c == '#')
                {
                    count++;
                }
                else
                {
                    break;
                }
            }

            if (count <= 0)
            {
                return 1;
            }

            return Math.Min(count, 9);
        }

        private static int ParseHeadingLevelToken(string levelToken)
        {
            if (string.IsNullOrWhiteSpace(levelToken))
            {
                return 0;
            }

            Match m = Regex.Match(levelToken.Trim(), @"(\d+)");
            if (m.Success && int.TryParse(m.Groups[1].Value, out int level))
            {
                return Math.Min(Math.Max(level, 1), 9);
            }

            return 0;
        }

        private static string MapHeadingLevelToRole(int level)
        {
            if (level <= 0)
            {
                return "body";
            }

            level = Math.Min(level, 9);
            return $"heading_l{level}";
        }
    }
}
