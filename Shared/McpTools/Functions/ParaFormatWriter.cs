using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 段落格式写入。只写本次传入的键。单键失败记入 format_warnings，不中断其余键。
    /// </summary>
    public static class ParaFormatWriter
    {
        public const string WarningStyleDirectOverride = "para_format_style_direct_override";

        public static Dictionary<string, object> SanitizeParaFormatForApply(
            IReadOnlyDictionary<string, object> paraFormat)
        {
            return ParaFormatCatalog.PrepareWriteSnapshot(paraFormat, new List<Dictionary<string, object>>());
        }

        public static List<string> Apply(
            Word.Range paragraphRange,
            IReadOnlyDictionary<string, object> paraFormat,
            out List<string> warnings,
            out List<Dictionary<string, object>> formatWarnings)
        {
            warnings = new List<string>();
            formatWarnings = new List<Dictionary<string, object>>();
            var applied = new List<string>();
            Dictionary<string, object> snap = ParaFormatCatalog.PrepareWriteSnapshot(paraFormat, formatWarnings);
            if (paragraphRange == null || snap.Count == 0)
            {
                return applied;
            }

            if (snap.ContainsKey("paragraph_style"))
            {
                TryWrite(applied, formatWarnings, "paragraph_style", () =>
                    WriteStyle(paragraphRange, snap["paragraph_style"]));
            }

            TryBool(paragraphRange, snap, applied, formatWarnings, "space_before_auto", (pf, v) => pf.SpaceBeforeAuto = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "space_after_auto", (pf, v) => pf.SpaceAfterAuto = v);

            TryAlignment(paragraphRange, snap, applied, formatWarnings);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "left_indent", (pf, v) => pf.LeftIndent = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "right_indent", (pf, v) => pf.RightIndent = v);
            TryFirstOrHanging(paragraphRange, snap, applied, formatWarnings);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "space_before", (pf, v) => pf.SpaceBefore = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "space_after", (pf, v) => pf.SpaceAfter = v);
            TryLineSpacing(paragraphRange, snap, applied, formatWarnings);

            TryFloat(paragraphRange, snap, applied, formatWarnings, "character_unit_left_indent", (pf, v) => pf.CharacterUnitLeftIndent = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "character_unit_right_indent", (pf, v) => pf.CharacterUnitRightIndent = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "character_unit_first_line_indent", (pf, v) => pf.CharacterUnitFirstLineIndent = v);

            TryEnum<Word.WdOutlineLevel>(paragraphRange, snap, applied, formatWarnings, "outline_level", (pf, v) => pf.OutlineLevel = v);
            TryEnum<Word.WdBaselineAlignment>(paragraphRange, snap, applied, formatWarnings, "baseline_alignment", (pf, v) => pf.BaseLineAlignment = v);
            TryEnum<Word.WdReadingOrder>(paragraphRange, snap, applied, formatWarnings, "reading_order", (pf, v) => pf.ReadingOrder = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "mirror_indents", (pf, v) => pf.MirrorIndents = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "auto_adjust_right_indent", (pf, v) => pf.AutoAdjustRightIndent = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "line_unit_before", (pf, v) => pf.LineUnitBefore = v);
            TryFloat(paragraphRange, snap, applied, formatWarnings, "line_unit_after", (pf, v) => pf.LineUnitAfter = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "disable_line_height_grid", (pf, v) => pf.DisableLineHeightGrid = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "widow_control", (pf, v) => pf.WidowControl = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "keep_with_next", (pf, v) => pf.KeepWithNext = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "keep_together", (pf, v) => pf.KeepTogether = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "page_break_before", (pf, v) => pf.PageBreakBefore = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "no_line_number", (pf, v) => pf.NoLineNumber = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "hyphenation", (pf, v) => pf.Hyphenation = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "far_east_line_break_control", (pf, v) => pf.FarEastLineBreakControl = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "word_wrap", (pf, v) => pf.WordWrap = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "hanging_punctuation", (pf, v) => pf.HangingPunctuation = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "half_width_punctuation_on_top_of_line", (pf, v) => pf.HalfWidthPunctuationOnTopOfLine = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "add_space_between_far_east_and_alpha", (pf, v) => WriteFarEastSpacing(paragraphRange, pf, v, true));
            TryBool(paragraphRange, snap, applied, formatWarnings, "add_space_between_far_east_and_digit", (pf, v) => WriteFarEastSpacing(paragraphRange, pf, v, false));
            TryEnum<Word.WdTextboxTightWrap>(paragraphRange, snap, applied, formatWarnings, "textbox_tight_wrap", (pf, v) => pf.TextboxTightWrap = v);
            TryBool(paragraphRange, snap, applied, formatWarnings, "collapsed_by_default", (pf, v) => pf.CollapsedByDefault = v);

            if (snap.ContainsKey("shading"))
            {
                TryWrite(applied, formatWarnings, "shading", () => WriteShading(paragraphRange, snap["shading"]));
            }

            if (snap.ContainsKey("borders"))
            {
                TryWrite(applied, formatWarnings, "borders", () => WriteBorders(paragraphRange, snap["borders"]));
            }

            if (snap.ContainsKey("tab_stops"))
            {
                TryWrite(applied, formatWarnings, "tab_stops", () => WriteTabStops(paragraphRange, snap["tab_stops"]));
            }

            if (snap.ContainsKey("list"))
            {
                TryWrite(applied, formatWarnings, "list", () => WriteList(paragraphRange, snap["list"]));
            }

            if (snap.ContainsKey("drop_cap"))
            {
                TryWrite(applied, formatWarnings, "drop_cap", () => WriteDropCap(paragraphRange, snap["drop_cap"]));
            }

            return applied;
        }

        private static void TryAlignment(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings)
        {
            if (!snap.ContainsKey("alignment"))
            {
                return;
            }

            TryWrite(applied, warnings, "alignment", () =>
            {
                range.ParagraphFormat.Alignment = ParseAlignment(snap["alignment"]?.ToString());
            });
        }

        private static void TryFirstOrHanging(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings)
        {
            if (snap.ContainsKey("hanging_indent"))
            {
                TryWrite(applied, warnings, "hanging_indent", () =>
                {
                    range.ParagraphFormat.FirstLineIndent = -ToFloat(snap["hanging_indent"]);
                });
                return;
            }

            if (snap.ContainsKey("first_line_indent"))
            {
                TryWrite(applied, warnings, "first_line_indent", () =>
                {
                    range.ParagraphFormat.FirstLineIndent = ToFloat(snap["first_line_indent"]);
                });
            }
        }

        private static void TryLineSpacing(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings)
        {
            if (!snap.ContainsKey("line_spacing_rule") && !snap.ContainsKey("line_spacing"))
            {
                return;
            }

            try
            {
                string rule = snap.TryGetValue("line_spacing_rule", out object r)
                    ? r?.ToString() ?? "single"
                    : "single";
                double spacing = snap.TryGetValue("line_spacing", out object s)
                    ? Convert.ToDouble(s, CultureInfo.InvariantCulture)
                    : 1.0;
                Dictionary<string, object> normalized = ParaFormatReader.NormalizeLineSpacingPair(rule, spacing);
                ApplyLineSpacingPair(
                    range.ParagraphFormat,
                    normalized["line_spacing_rule"]?.ToString(),
                    Convert.ToDouble(normalized["line_spacing"], CultureInfo.InvariantCulture));
                if (snap.ContainsKey("line_spacing_rule"))
                {
                    applied.Add("line_spacing_rule");
                }

                if (snap.ContainsKey("line_spacing"))
                {
                    applied.Add("line_spacing");
                }
            }
            catch (Exception ex)
            {
                Warn(warnings, "line_spacing_rule", ex.Message);
            }
        }

        private static void TryFloat(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings,
            string key,
            Action<Word.ParagraphFormat, float> write)
        {
            if (!snap.ContainsKey(key))
            {
                return;
            }

            TryWrite(applied, warnings, key, () => write(range.ParagraphFormat, ToFloat(snap[key])));
        }

        private static void TryBool(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings,
            string key,
            Action<Word.ParagraphFormat, int> write)
        {
            if (!snap.ContainsKey(key))
            {
                return;
            }

            TryWrite(applied, warnings, key, () => write(range.ParagraphFormat, ToBoolInt(snap[key])));
        }

        private static void TryEnum<T>(
            Word.Range range,
            Dictionary<string, object> snap,
            List<string> applied,
            List<Dictionary<string, object>> warnings,
            string key,
            Action<Word.ParagraphFormat, T> write)
            where T : struct
        {
            if (!snap.ContainsKey(key))
            {
                return;
            }

            TryWrite(applied, warnings, key, () => write(range.ParagraphFormat, ParseEnum<T>(snap[key])));
        }

        private static void TryWrite(
            List<string> applied,
            List<Dictionary<string, object>> warnings,
            string key,
            Action write)
        {
            try
            {
                write();
                applied.Add(key);
            }
            catch (Exception ex)
            {
                Warn(warnings, key, ex.Message);
            }
        }

        private static void WriteStyle(Word.Range range, object value)
        {
            string name = value?.ToString()?.Trim() ?? "";
            if (name.Length == 0)
            {
                throw new InvalidOperationException("样式名为空");
            }

            object styleRef = name;
            range.set_Style(ref styleRef);
        }

        private static void WriteShading(Word.Range range, object raw)
        {
            Dictionary<string, object> shading = AsDict(raw);
            if (shading == null)
            {
                throw new InvalidOperationException("shading 须为对象");
            }

            Word.Shading target = range.ParagraphFormat.Shading;
            if (shading.ContainsKey("texture"))
            {
                target.Texture = ParseEnum<Word.WdTextureIndex>(shading["texture"]);
            }

            if (shading.ContainsKey("background_pattern_color"))
            {
                target.BackgroundPatternColor = ParseColor(shading["background_pattern_color"]);
            }

            if (shading.ContainsKey("foreground_pattern_color"))
            {
                target.ForegroundPatternColor = ParseColor(shading["foreground_pattern_color"]);
            }
        }

        private static void WriteBorders(Word.Range range, object raw)
        {
            Dictionary<string, object> borders = AsDict(raw);
            if (borders == null)
            {
                throw new InvalidOperationException("borders 须为对象");
            }

            Word.Borders target = range.ParagraphFormat.Borders;
            WriteBorder(target, borders, "top", Word.WdBorderType.wdBorderTop);
            WriteBorder(target, borders, "bottom", Word.WdBorderType.wdBorderBottom);
            WriteBorder(target, borders, "left", Word.WdBorderType.wdBorderLeft);
            WriteBorder(target, borders, "right", Word.WdBorderType.wdBorderRight);
            WriteBorder(target, borders, "horizontal", Word.WdBorderType.wdBorderHorizontal);
        }

        private static void WriteBorder(
            Word.Borders borders,
            Dictionary<string, object> source,
            string key,
            Word.WdBorderType which)
        {
            if (!source.ContainsKey(key))
            {
                return;
            }

            Dictionary<string, object> edge = AsDict(source[key]);
            if (edge == null)
            {
                throw new InvalidOperationException(key + " 须为对象");
            }

            Word.Border border = borders[which];
            if (edge.ContainsKey("line_style"))
            {
                border.LineStyle = ParseEnum<Word.WdLineStyle>(edge["line_style"]);
            }

            if (edge.ContainsKey("line_width"))
            {
                border.LineWidth = ParseEnum<Word.WdLineWidth>(edge["line_width"]);
            }

            if (edge.ContainsKey("color"))
            {
                border.Color = ParseColor(edge["color"]);
            }
        }

        private static void WriteTabStops(Word.Range range, object raw)
        {
            if (raw is string || raw is IDictionary)
            {
                throw new InvalidOperationException("tab_stops 须为数组");
            }

            Word.TabStops stops = range.ParagraphFormat.TabStops;
            for (int i = stops.Count; i >= 1; i--)
            {
                object index = i;
                Word.TabStop stop = stops.get_Item(ref index);
                if (stop.CustomTab)
                {
                    stop.Clear();
                }
            }

            if (!(raw is IEnumerable seq))
            {
                return;
            }

            foreach (object item in seq)
            {
                Dictionary<string, object> tab = AsDict(item);
                if (tab == null || !tab.ContainsKey("position"))
                {
                    throw new InvalidOperationException("制表位缺少 position");
                }

                float position = ToFloat(tab["position"]);
                object alignment = tab.ContainsKey("alignment")
                    ? (object)ParseEnum<Word.WdTabAlignment>(tab["alignment"])
                    : Word.WdTabAlignment.wdAlignTabLeft;
                object leader = tab.ContainsKey("leader")
                    ? (object)ParseEnum<Word.WdTabLeader>(tab["leader"])
                    : Word.WdTabLeader.wdTabLeaderSpaces;
                stops.Add(position, ref alignment, ref leader);
            }
        }

        private static void WriteList(Word.Range range, object raw)
        {
            Dictionary<string, object> list = AsDict(raw);
            if (list == null)
            {
                throw new InvalidOperationException("list 须为对象");
            }

            Word.ListFormat listFormat = range.ListFormat;
            string kind = list.TryGetValue("kind", out object kindObj)
                ? (kindObj?.ToString() ?? "").Trim().ToLowerInvariant()
                : "";
            bool restart = list.ContainsKey("continue_previous") && !IsTrue(list["continue_previous"]);
            bool join = list.ContainsKey("continue_previous") && IsTrue(list["continue_previous"]);
            if (kind.Length > 0)
            {
                if (!(join && IsWpsHost(range) && TryJoinPreviousList(range, listFormat)))
                {
                    ApplyListKind(listFormat, kind, list);
                }
            }

            if (restart)
            {
                RestartList(listFormat);
            }

            if (list.ContainsKey("level"))
            {
                listFormat.ListLevelNumber = Convert.ToInt32(list["level"], CultureInfo.InvariantCulture);
            }

            bool hasLook = list.ContainsKey("number_style")
                || list.ContainsKey("number_format")
                || list.ContainsKey("start_at")
                || list.ContainsKey("bullet_char")
                || list.ContainsKey("picture_path");
            if (!hasLook)
            {
                return;
            }

            if (kind == "none" || listFormat.ListTemplate == null)
            {
                throw new InvalidOperationException("这段还没有列表，不能写编号样子");
            }

            if (IsWpsHost(range))
            {
                ApplyWpsListLook(range, listFormat, list, kind, restart);
                return;
            }

            int level = listFormat.ListLevelNumber;
            if (level < 1)
            {
                level = 1;
            }

            Word.ListLevel listLevel = listFormat.ListTemplate.ListLevels[level];
            if (list.ContainsKey("bullet_char"))
            {
                listLevel.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleBullet;
                listLevel.NumberFormat = list["bullet_char"]?.ToString() ?? "";
            }

            if (list.ContainsKey("number_style"))
            {
                listLevel.NumberStyle = ParseEnum<Word.WdListNumberStyle>(list["number_style"]);
            }

            if (list.ContainsKey("number_format"))
            {
                listLevel.NumberFormat = list["number_format"]?.ToString() ?? "";
            }

            if (list.ContainsKey("start_at"))
            {
                listLevel.StartAt = Convert.ToInt32(list["start_at"], CultureInfo.InvariantCulture);
            }

            if (list.ContainsKey("picture_path") && kind != "picture_bullet")
            {
                if (!TryApplyPictureBullet(listFormat, listLevel, list["picture_path"]?.ToString()))
                {
                    throw new InvalidOperationException("图片项目符号没有写上");
                }
            }
        }

        private static void ApplyListKind(Word.ListFormat listFormat, string kind, Dictionary<string, object> list)
        {
            object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
            switch (kind)
            {
                case "none":
                    object numberType = Type.Missing;
                    listFormat.RemoveNumbers(ref numberType);
                    break;
                case "bullet":
                    listFormat.ApplyBulletDefault(ref behavior);
                    break;
                case "number":
                    listFormat.ApplyNumberDefault(ref behavior);
                    break;
                case "outline":
                    listFormat.ApplyOutlineNumberDefault(ref behavior);
                    break;
                case "picture_bullet":
                    string path = list.TryGetValue("picture_path", out object pathObj)
                        ? pathObj?.ToString() ?? ""
                        : "";
                    if (path.Length == 0)
                    {
                        throw new InvalidOperationException("picture_bullet 缺少 picture_path");
                    }

                    listFormat.ApplyBulletDefault(ref behavior);
                    int level = listFormat.ListLevelNumber < 1 ? 1 : listFormat.ListLevelNumber;
                    Word.Range owner = listFormat.Parent as Word.Range;
                    if (!TryApplyPictureBullet(listFormat, listFormat.ListTemplate.ListLevels[level], path)
                        && (owner == null || !IsWpsHost(owner)))
                    {
                        throw new InvalidOperationException("图片项目符号没有写上");
                    }

                    break;
                default:
                    throw new InvalidOperationException("无法识别的 list.kind " + kind);
            }
        }

        private static void RestartList(Word.ListFormat listFormat)
        {
            Word.ListTemplate template = listFormat.ListTemplate;
            if (template == null)
            {
                throw new InvalidOperationException("没有可另起的列表");
            }

            object continuePrevious = false;
            object applyTo = Word.WdListApplyTo.wdListApplyToSelection;
            object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
            listFormat.ApplyListTemplate(template, ref continuePrevious, ref applyTo, ref behavior);
        }

        private static void ApplyWpsListLook(
            Word.Range range,
            Word.ListFormat listFormat,
            Dictionary<string, object> list,
            string kind,
            bool restart)
        {
            int level = listFormat.ListLevelNumber;
            if (level < 1)
            {
                level = 1;
            }

            object outline = string.Equals(kind, "outline", StringComparison.Ordinal);
            object name = Type.Missing;
            Word.ListTemplate fresh = range.Document.ListTemplates.Add(ref outline, ref name);
            Word.ListLevel src = listFormat.ListTemplate.ListLevels[level];
            Word.ListLevel dst = fresh.ListLevels[level];
            CopyListLevel(src, dst);
            if (list.ContainsKey("bullet_char"))
            {
                dst.NumberStyle = Word.WdListNumberStyle.wdListNumberStyleBullet;
                dst.NumberFormat = list["bullet_char"]?.ToString() ?? "";
            }

            if (list.ContainsKey("number_style"))
            {
                dst.NumberStyle = ParseEnum<Word.WdListNumberStyle>(list["number_style"]);
            }

            if (list.ContainsKey("number_format"))
            {
                dst.NumberFormat = list["number_format"]?.ToString() ?? "";
            }

            if (list.ContainsKey("start_at"))
            {
                dst.StartAt = Convert.ToInt32(list["start_at"], CultureInfo.InvariantCulture);
            }

            object cont = !restart;
            object applyTo = Word.WdListApplyTo.wdListApplyToWholeList;
            object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
            listFormat.ApplyListTemplate(fresh, ref cont, ref applyTo, ref behavior);
            if (list.ContainsKey("picture_path"))
            {
                int appliedLevel = listFormat.ListLevelNumber < 1 ? 1 : listFormat.ListLevelNumber;
                if (!TryApplyPictureBullet(listFormat, listFormat.ListTemplate.ListLevels[appliedLevel], list["picture_path"]?.ToString()))
                {
                    throw new InvalidOperationException("图片项目符号没有写上");
                }
            }
        }

        private static void CopyListLevel(Word.ListLevel src, Word.ListLevel dst)
        {
            try { dst.NumberStyle = src.NumberStyle; } catch { }
            try { dst.NumberFormat = src.NumberFormat; } catch { }
            try { dst.StartAt = src.StartAt; } catch { }
            try { dst.Alignment = src.Alignment; } catch { }
            try { dst.TrailingCharacter = src.TrailingCharacter; } catch { }
            try { dst.NumberPosition = src.NumberPosition; } catch { }
            try { dst.TextPosition = src.TextPosition; } catch { }
            try { dst.TabPosition = src.TabPosition; } catch { }
            try { dst.ResetOnHigher = src.ResetOnHigher; } catch { }
            try { dst.Font.Name = src.Font.Name; } catch { }
        }

        private static bool IsWpsHost(Word.Range range)
        {
            try
            {
                string name = range.Application?.Name ?? "";
                return name.IndexOf("WPS", StringComparison.OrdinalIgnoreCase) >= 0
                    || name.IndexOf("Kingsoft", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>接到上一段的同一份列表。WPS 上 ApplyNumberDefault 会另起一份。</summary>
        private static bool TryJoinPreviousList(Word.Range range, Word.ListFormat listFormat)
        {
            if (range.Paragraphs == null || range.Paragraphs.Count < 1)
            {
                return false;
            }

            Word.Paragraph prev = range.Paragraphs[1].Previous();
            for (int i = 0; i < 20 && prev != null; i++)
            {
                Word.ListFormat prevList = prev.Range.ListFormat;
                if (prevList != null
                    && prevList.ListTemplate != null
                    && prevList.ListType != Word.WdListType.wdListNoNumbering)
                {
                    object cont = true;
                    object applyTo = Word.WdListApplyTo.wdListApplyToSelection;
                    object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
                    listFormat.ApplyListTemplate(prevList.ListTemplate, ref cont, ref applyTo, ref behavior);
                    return true;
                }

                prev = prev.Previous();
            }

            return false;
        }

        /// <summary>改完列表级后再套一次，WPS 才会把编号样子写进段落。</summary>
        private static void ReapplyListTemplate(Word.ListFormat listFormat, bool continuePrevious, bool wholeList)
        {
            Word.ListTemplate template = listFormat.ListTemplate;
            if (template == null)
            {
                return;
            }

            object cont = continuePrevious;
            object applyTo = wholeList
                ? (object)Word.WdListApplyTo.wdListApplyToWholeList
                : Word.WdListApplyTo.wdListApplyToSelection;
            object behavior = Word.WdDefaultListBehavior.wdWord10ListBehavior;
            listFormat.ApplyListTemplate(template, ref cont, ref applyTo, ref behavior);
        }

        private static bool TryApplyPictureBullet(Word.ListFormat listFormat, Word.ListLevel listLevel, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                throw new InvalidOperationException("picture_path 为空");
            }

            try
            {
                listLevel.ApplyPictureBullet(path);
            }
            catch
            {
                /* 换 bmp 再试 */
            }

            if (listFormat.ListType == Word.WdListType.wdListPictureBullet)
            {
                return true;
            }

            string bmp = null;
            try
            {
                bmp = EnsureBmp(path);
            }
            catch
            {
                return false;
            }

            if (!string.Equals(bmp, path, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    listLevel.ApplyPictureBullet(bmp);
                }
                catch
                {
                    return listFormat.ListType == Word.WdListType.wdListPictureBullet;
                }
            }

            return listFormat.ListType == Word.WdListType.wdListPictureBullet;
        }

        private static string EnsureBmp(string path)
        {
            if (path.EndsWith(".bmp", StringComparison.OrdinalIgnoreCase))
            {
                return path;
            }

            string bmp = Path.Combine(Path.GetTempPath(), "ew-pf-" + Guid.NewGuid().ToString("N") + ".bmp");
            using (Image image = Image.FromFile(path))
            {
                image.Save(bmp, ImageFormat.Bmp);
            }

            return bmp;
        }

        private static void WriteFarEastSpacing(Word.Range range, Word.ParagraphFormat pf, int value, bool alpha)
        {
            if (alpha)
            {
                pf.AddSpaceBetweenFarEastAndAlpha = value;
            }
            else
            {
                pf.AddSpaceBetweenFarEastAndDigit = value;
            }

            if (FarEastSpacingMatches(pf, value, alpha))
            {
                return;
            }

            string name = alpha ? "AddSpaceBetweenFarEastAndAlpha" : "AddSpaceBetweenFarEastAndDigit";
            pf.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty,
                null,
                pf,
                new object[] { value != 0 });
            if (FarEastSpacingMatches(pf, value, alpha))
            {
                return;
            }

            PatchFarEastSpacingXml(range, alpha, value != 0);
            if (FarEastSpacingMatches(pf, value, alpha))
            {
                return;
            }

            string tag = alpha ? "autoSpaceDE" : "autoSpaceDN";
            string xml = "";
            try
            {
                xml = range.Paragraphs[1].Range.WordOpenXML ?? "";
            }
            catch
            {
                xml = "";
            }

            if (!XmlBoolIs(xml, tag, value != 0))
            {
                throw new InvalidOperationException(name + " 写上后读回不是预期值");
            }
        }

        private static bool XmlBoolIs(string xml, string tag, bool on)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return false;
            }

            Match match = Regex.Match(xml, "<w:" + tag + "\\b[^>]*/>");
            if (!match.Success)
            {
                return false;
            }

            bool isOff = match.Value.IndexOf("w:val=\"0\"", StringComparison.Ordinal) >= 0;
            return on ? !isOff : isOff;
        }

        private static bool FarEastSpacingMatches(Word.ParagraphFormat pf, int value, bool alpha)
        {
            int got = alpha ? pf.AddSpaceBetweenFarEastAndAlpha : pf.AddSpaceBetweenFarEastAndDigit;
            if (got == ParaFormatCatalog.WdUndefined || got == ParaFormatCatalog.WdToggle)
            {
                return false;
            }

            return (got != 0) == (value != 0);
        }

        private static void PatchFarEastSpacingXml(Word.Range range, bool alpha, bool on)
        {
            if (range.Paragraphs == null || range.Paragraphs.Count < 1)
            {
                throw new InvalidOperationException("没有段落可写中西文间距");
            }

            Word.Range para = range.Paragraphs[1].Range;
            string xml = para.WordOpenXML;
            if (string.IsNullOrEmpty(xml))
            {
                throw new InvalidOperationException("段落 XML 为空");
            }

            string tag = alpha ? "autoSpaceDE" : "autoSpaceDN";
            string patched = UpsertPPrBool(xml, tag, on ? "1" : "0");
            para.InsertXML(patched);
        }

        private static string UpsertPPrBool(string xml, string tag, string val)
        {
            string element = "<w:" + tag + " w:val=\"" + val + "\"/>";
            var existing = new Regex("<w:" + tag + "\\b[^>]*/>");
            if (existing.IsMatch(xml))
            {
                return existing.Replace(xml, element, 1);
            }

            int close = xml.IndexOf("</w:pPr>", StringComparison.Ordinal);
            if (close >= 0)
            {
                return xml.Insert(close, element);
            }

            int open = xml.IndexOf("<w:p", StringComparison.Ordinal);
            if (open < 0)
            {
                throw new InvalidOperationException("段落 XML 里没有 w:p");
            }

            int gt = xml.IndexOf('>', open);
            if (gt < 0)
            {
                throw new InvalidOperationException("段落 XML 的 w:p 不完整");
            }

            return xml.Insert(gt + 1, "<w:pPr>" + element + "</w:pPr>");
        }

        private static void WriteDropCap(Word.Range range, object raw)
        {
            Dictionary<string, object> drop = AsDict(raw);
            if (drop == null)
            {
                throw new InvalidOperationException("drop_cap 须为对象");
            }

            if (range.Paragraphs == null || range.Paragraphs.Count < 1)
            {
                throw new InvalidOperationException("没有段落可写首字下沉");
            }

            foreach (Word.Paragraph paragraph in range.Paragraphs)
            {
                Word.DropCap target = paragraph.DropCap;
                if (drop.ContainsKey("position"))
                {
                    target.Position = ParseDropPosition(drop["position"]?.ToString());
                }

                if (drop.ContainsKey("lines"))
                {
                    target.LinesToDrop = Convert.ToInt32(drop["lines"], CultureInfo.InvariantCulture);
                }

                if (drop.ContainsKey("font_name"))
                {
                    target.FontName = drop["font_name"]?.ToString() ?? "";
                }

                if (drop.ContainsKey("distance"))
                {
                    target.DistanceFromText = ToFloat(drop["distance"]);
                }
            }
        }

        internal static void ApplyLineSpacingPair(Word.ParagraphFormat pf, string rule, double spacing)
        {
            switch ((rule ?? "single").ToLowerInvariant())
            {
                case "1.5":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    break;
                case "double":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                    break;
                case "at_least":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;
                    pf.LineSpacing = (float)spacing;
                    break;
                case "exact":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    pf.LineSpacing = (float)spacing;
                    break;
                case "multiple":
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    pf.LineSpacing = spacing <= 3.0 ? (float)spacing * 12f : (float)spacing;
                    break;
                default:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    break;
            }
        }

        private static Word.WdParagraphAlignment ParseAlignment(string s)
        {
            switch ((s ?? "").ToLowerInvariant())
            {
                case "center":
                    return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case "right":
                    return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case "justify":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                case "distribute":
                    return Word.WdParagraphAlignment.wdAlignParagraphDistribute;
                case "justify_med":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustifyMed;
                case "justify_hi":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustifyHi;
                case "justify_low":
                    return Word.WdParagraphAlignment.wdAlignParagraphJustifyLow;
                case "thai_justify":
                    return Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;
                case "left":
                    return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                default:
                    throw new InvalidOperationException("无法识别的对齐 " + s);
            }
        }

        private static Word.WdDropPosition ParseDropPosition(string s)
        {
            switch ((s ?? "").ToLowerInvariant())
            {
                case "drop":
                    return Word.WdDropPosition.wdDropNormal;
                case "margin":
                    return Word.WdDropPosition.wdDropMargin;
                case "none":
                    return Word.WdDropPosition.wdDropNone;
                default:
                    throw new InvalidOperationException("无法识别的首字下沉位置 " + s);
            }
        }

        private static T ParseEnum<T>(object value)
            where T : struct
        {
            string text = value?.ToString()?.Trim() ?? "";
            if (!Enum.TryParse(text, true, out T parsed))
            {
                throw new InvalidOperationException("无法识别 " + text);
            }

            return parsed;
        }

        private static Word.WdColor ParseColor(object value)
        {
            string rgb = value?.ToString()?.Trim() ?? "";
            if (rgb.Length == 0 || string.Equals(rgb, "automatic", StringComparison.OrdinalIgnoreCase))
            {
                return Word.WdColor.wdColorAutomatic;
            }

            string hex = rgb.TrimStart('#');
            if (hex.Length != 6)
            {
                throw new InvalidOperationException("颜色须为 #RRGGBB 或 automatic");
            }

            int r = Convert.ToInt32(hex.Substring(0, 2), 16);
            int g = Convert.ToInt32(hex.Substring(2, 2), 16);
            int b = Convert.ToInt32(hex.Substring(4, 2), 16);
            return (Word.WdColor)((b << 16) | (g << 8) | r);
        }

        private static float ToFloat(object value)
        {
            if (ParaFormatCatalog.IsSentinel(value))
            {
                throw new InvalidOperationException("哨兵值不写入");
            }

            return Convert.ToSingle(value, CultureInfo.InvariantCulture);
        }

        private static int ToBoolInt(object value)
        {
            if (ParaFormatCatalog.IsSentinel(value))
            {
                throw new InvalidOperationException("哨兵值不写入");
            }

            if (value is bool flag)
            {
                return flag ? -1 : 0;
            }

            string text = value?.ToString()?.Trim() ?? "";
            if (string.Equals(text, "true", StringComparison.OrdinalIgnoreCase))
            {
                return -1;
            }

            if (string.Equals(text, "false", StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            int number = Convert.ToInt32(value, CultureInfo.InvariantCulture);
            return number != 0 ? -1 : 0;
        }

        private static bool IsTrue(object value)
        {
            if (value is bool flag)
            {
                return flag;
            }

            string text = value?.ToString()?.Trim() ?? "";
            return text == "1" || text == "-1" || string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }

        private static Dictionary<string, object> AsDict(object raw)
        {
            if (raw is Dictionary<string, object> dict)
            {
                return dict;
            }

            if (raw is IDictionary generic)
            {
                var copy = new Dictionary<string, object>(StringComparer.Ordinal);
                foreach (DictionaryEntry entry in generic)
                {
                    copy[Convert.ToString(entry.Key)] = entry.Value;
                }

                return copy;
            }

            return null;
        }

        private static void Warn(List<Dictionary<string, object>> warnings, string key, string message)
        {
            warnings.Add(new Dictionary<string, object>
            {
                ["key"] = key ?? "",
                ["message"] = message ?? "",
            });
        }
    }
}
