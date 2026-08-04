using System;
using System.Collections.Generic;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// replace ≤2 自动 inherit：从 Range 提取/应用方案 B 格式快照（Phase 1）。
    /// </summary>
    public static class FormatInheritHelper
    {
        /// <summary>调试日志统一前缀；写入 WordAddIn2/logs/ 当前会话文件，VS 输出窗口可搜：<c>[格式融合][DBG]</c></summary>
        public const string DbgPrefix = "[格式融合][DBG]";

        /// <summary>
        /// 格式融合 / BodyFormat 明细 DBG 总开关。false = 代码保留但不输出；排查时改 true。
        /// </summary>
        /// <summary>调试开关；用 static readonly（非 const）避免 EnableDbg=false 时 CS0162 不可达警告。</summary>
        public static readonly bool EnableDbg = false;

        public static Dictionary<string, object> ExtractSnapshot(Word.Range range)
        {
            var format = new Dictionary<string, object>();
            if (range == null)
            {
                DbgLog("ExtractSnapshot range=null");
                return format;
            }

            DbgLogRangeFonts("ExtractSnapshot BEFORE", range);

            try
            {
                format["font_name"] = GetEffectiveFontName(range.Font);
                format["font_size"] = range.Font?.Size ?? 0f;
                format["font_color"] = ConvertWdColorToRgb(range.Font?.Color ?? Word.WdColor.wdColorAutomatic);
                format["background_color"] = ConvertWdColorToRgb(
                    range.Shading?.BackgroundPatternColor ?? Word.WdColor.wdColorAutomatic);
                format["is_bold"] = (range.Font?.Bold ?? 0) != 0;
                format["has_underline"] = (range.Font?.Underline ?? Word.WdUnderline.wdUnderlineNone)
                    != Word.WdUnderline.wdUnderlineNone;
                format["has_strikethrough"] = (range.Font?.StrikeThrough ?? 0) != 0;

                try
                {
                    var style = range.get_Style() as Word.Style;
                    format["paragraph_style"] = style?.NameLocal ?? "";
                }
                catch (Exception ex)
                {
                    DbgLog($"ExtractSnapshot paragraph_style 读取失败: {ex.Message}");
                    format["paragraph_style"] = "";
                }

                format["alignment"] = MapAlignment(range.ParagraphFormat?.Alignment ?? Word.WdParagraphAlignment.wdAlignParagraphLeft);

                DbgLog($"ExtractSnapshot snapshot={SnapshotToDebugString(format)}");
            }
            catch (Exception ex)
            {
                DbgLog($"ExtractSnapshot 异常: {ex.Message}");
            }

            return format;
        }

        /// <param name="charFormatOnly">为 true 时只写字体等字符格式，不 set_Style / 段落对齐（避免 partial range 触发整段样式重置为等线）</param>
        public static void ApplySnapshot(Word.Range range, Dictionary<string, object> snapshot, bool charFormatOnly = false)
        {
            if (range == null || snapshot == null || snapshot.Count == 0)
            {
                DbgLog("ApplySnapshot 跳过 range/snapshot 为空");
                return;
            }

            DbgLogRangeFonts("ApplySnapshot BEFORE", range);
            DbgLog($"ApplySnapshot snapshot={SnapshotToDebugString(snapshot)} charFormatOnly={charFormatOnly}");

            if (!charFormatOnly
                && snapshot.TryGetValue("paragraph_style", out object ps)
                && ps is string styleName && !string.IsNullOrEmpty(styleName))
            {
                try
                {
                    DbgLog($"ApplySnapshot set_Style \"{styleName}\"");
                    range.set_Style(styleName);
                    DbgLogRangeFonts("ApplySnapshot AFTER set_Style", range);
                }
                catch (Exception ex)
                {
                    DbgLog($"ApplySnapshot set_Style 失败: {ex.Message}");
                }
            }

            ApplyCharFormat(range, snapshot);
            DbgLogRangeFonts("ApplySnapshot AFTER char", range);

            if (!charFormatOnly
                && snapshot.TryGetValue("alignment", out object al) && al is string alignStr)
            {
                range.ParagraphFormat.Alignment = ParseAlignment(alignStr);
                DbgLog($"ApplySnapshot alignment={alignStr}");
            }

            DbgLogRangeFonts("ApplySnapshot AFTER", range);
        }

        public static string BuildFingerprint(Dictionary<string, object> snapshot)
        {
            if (snapshot == null)
            {
                return "";
            }

            string font = snapshot.TryGetValue("font_name", out object fn) ? fn?.ToString() ?? "" : "";
            string size = snapshot.TryGetValue("font_size", out object fs) ? fs?.ToString() ?? "" : "";
            string style = snapshot.TryGetValue("paragraph_style", out object ps) ? ps?.ToString() ?? "" : "";
            bool bold = snapshot.TryGetValue("is_bold", out object b) && b is bool bb && bb;
            return $"{font}|{size}|{style}|{(bold ? "加粗" : "非加粗")}";
        }

        /// <summary>
        /// 从 explicit char_format 参数构建快照（仅包含传入字段；Phase 3 不写入 paragraph_style / alignment）。
        /// </summary>
        /// <param name="includeBackgroundColor">false 时不写入 background_color（粗迁整篇 body 用，避免覆盖表格底纹等）</param>
        public static Dictionary<string, object> BuildSnapshotFromCharFormat(
            Dictionary<string, object> charFormat,
            bool includeBackgroundColor = true)
        {
            var snapshot = new Dictionary<string, object>();
            if (charFormat == null || charFormat.Count == 0)
            {
                return snapshot;
            }

            string[] allowed =
            {
                "font_name", "font_size", "font_color", "background_color",
                "is_bold", "has_underline", "has_strikethrough"
            };

            foreach (string key in allowed)
            {
                if (!includeBackgroundColor && key == "background_color")
                {
                    continue;
                }

                if (charFormat.TryGetValue(key, out object value) && value != null)
                {
                    snapshot[key] = value;
                }
            }

            return snapshot;
        }

        public static List<string> ListExplicitFieldsApplied(Dictionary<string, object> charFormat)
        {
            var applied = new List<string>();
            if (charFormat == null)
            {
                return applied;
            }

            foreach (string key in BuildSnapshotFromCharFormat(charFormat).Keys)
            {
                applied.Add(key);
            }

            return applied;
        }

        /// <summary>
        /// 将 <see cref="DocumentState.BodyFormat"/> 套到新写入句子（charFormatOnly）。
        /// Body 为空则跳过并返回 false。
        /// </summary>
        public static bool TryApplyBodyFormatToSentences(
            Word.Document doc,
            Word.Range blockRange,
            List<string> sentences,
            Func<Word.Document, Word.Range, string, Word.Range> findSentenceInRange)
        {
            if (!DocumentState.HasBodyFormat)
            {
                DbgLog("TryApplyBodyFormatToSentences 跳过：BodyFormat 为空");
                return false;
            }

            Dictionary<string, object> snapshot = DocumentState.BodyFormat;
            if (snapshot == null || snapshot.Count == 0)
            {
                DbgLog("TryApplyBodyFormatToSentences 跳过：BodyFormat 拷贝为空");
                return false;
            }

            try
            {
                DbgLog($"TryApplyBodyFormatToSentences 开始 ApplySnapshotToSentences BodyFormat={SnapshotToDebugString(snapshot)}");
                ApplySnapshotToSentences(doc, blockRange, sentences, snapshot, findSentenceInRange);
                DbgLog("TryApplyBodyFormatToSentences 完成");
                return true;
            }
            catch (Exception ex)
            {
                DbgLog($"TryApplyBodyFormatToSentences 失败: {ex.Message}");
                return false;
            }
        }

        public static void ApplySnapshotToSentences(
            Word.Document doc,
            Word.Range blockRange,
            List<string> sentences,
            Dictionary<string, object> snapshot,
            Func<Word.Document, Word.Range, string, Word.Range> findSentenceInRange)
        {
            if (doc == null || blockRange == null || sentences == null || snapshot == null || findSentenceInRange == null)
            {
                DbgLog("ApplySnapshotToSentences 跳过参数为空");
                return;
            }

            DbgLog($"ApplySnapshotToSentences block={blockRange.Start}-{blockRange.End} sentenceCount={sentences.Count}");

            int currentPos = blockRange.Start;
            int blockEnd = blockRange.End;
            int idx = 0;
            int appliedCount = 0;

            foreach (string sentence in sentences)
            {
                idx++;
                if (string.IsNullOrEmpty(sentence) || currentPos >= blockEnd)
                {
                    break;
                }

                Word.Range searchRange = doc.Range(currentPos, blockEnd);
                Word.Range sentenceRange = findSentenceInRange(doc, searchRange, sentence);
                if (sentenceRange == null)
                {
                    DbgLog($"ApplySnapshotToSentences 句{idx} 未找到 pos={currentPos} preview=\"{PreviewText(sentence)}\"");
                    currentPos += sentence.Length;
                    continue;
                }

                DbgLog($"ApplySnapshotToSentences 句{idx} found={sentenceRange.Start}-{sentenceRange.End} preview=\"{PreviewText(sentence)}\"");
                ApplySnapshot(sentenceRange, snapshot, charFormatOnly: true);
                appliedCount++;
                currentPos = sentenceRange.End;
            }

            if (appliedCount < sentences.Count)
            {
                DbgLog($"ApplySnapshotToSentences 回退：对整块 {blockRange.Start}-{blockRange.End} 应用字符格式（{appliedCount}/{sentences.Count} 句逐句成功）");
                ApplySnapshot(blockRange, snapshot, charFormatOnly: true);
            }
        }

        /// <summary>扫描文档前若干段字体（用于排查「整篇变等线」），搜索 <c>[格式融合][DBG][DOC-FONT]</c></summary>
        public static void DbgLogDocumentFontSample(Word.Document doc, string stage, int maxParagraphs = 8)
        {
            if (doc == null)
            {
                DbgLog("[DOC-FONT] doc=null stage=" + stage);
                return;
            }

            try
            {
                int total = doc.Paragraphs.Count;
                int n = Math.Min(maxParagraphs, total);
                DbgLog($"[DOC-FONT] stage={stage} paragraphs={total} sampleFirst={n}");
                for (int i = 1; i <= n; i++)
                {
                    Word.Range r = doc.Paragraphs[i].Range;
                    DbgLogRangeFonts($"[DOC-FONT] P{i}", r);
                }
            }
            catch (Exception ex)
            {
                DbgLog($"[DOC-FONT] stage={stage} 扫描失败: {ex.Message}");
            }
        }

        public static void DbgLog(string message)
        {
            if (!EnableDbg)
            {
                return;
            }

            System.Diagnostics.Debug.WriteLine($"{DbgPrefix} {message}");
        }

        public static void DbgLogRangeFonts(string tag, Word.Range range)
        {
            if (range == null)
            {
                DbgLog($"{tag} range=null");
                return;
            }

            try
            {
                string name = SafeFontString(() => range.Font?.Name);
                string farEast = SafeFontString(() => range.Font?.NameFarEast);
                string ascii = SafeFontString(() => range.Font?.NameAscii);
                string style = "";
                try
                {
                    style = ((Word.Style)range.get_Style())?.NameLocal ?? "";
                }
                catch
                {
                    style = "?";
                }

                float size = 0f;
                try { size = range.Font?.Size ?? 0f; } catch { }

                DbgLog(
                    $"{tag} pos={range.Start}-{range.End} len={range.End - range.Start} " +
                    $"Name=\"{name}\" NameFarEast=\"{farEast}\" NameAscii=\"{ascii}\" " +
                    $"Size={size} style=\"{style}\" text=\"{PreviewText(range.Text)}\"");
            }
            catch (Exception ex)
            {
                DbgLog($"{tag} 记录字体失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 快照调试串（inherit / BodyFormat 共用）。可搜日志：<c>snapshot=</c> 或 <c>mode=</c>。
        /// </summary>
        public static string SnapshotToDebugString(Dictionary<string, object> snapshot)
        {
            if (snapshot == null)
            {
                return "{}";
            }

            var sb = new StringBuilder("{");
            sb.Append($"font_name={Fmt(snapshot, "font_name")}");
            sb.Append($", font_size={Fmt(snapshot, "font_size")}");
            sb.Append($", font_color={Fmt(snapshot, "font_color")}");
            sb.Append($", background_color={Fmt(snapshot, "background_color")}");
            sb.Append($", is_bold={Fmt(snapshot, "is_bold")}");
            sb.Append($", has_underline={Fmt(snapshot, "has_underline")}");
            sb.Append($", has_strikethrough={Fmt(snapshot, "has_strikethrough")}");
            sb.Append($", paragraph_style={Fmt(snapshot, "paragraph_style")}");
            sb.Append($", alignment={Fmt(snapshot, "alignment")}");
            sb.Append("}");
            return sb.ToString();
        }

        private static string Fmt(Dictionary<string, object> d, string key)
        {
            return d.TryGetValue(key, out object v) ? v?.ToString() ?? "" : "";
        }

        private static string PreviewText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string t = text.Replace("\r", "\\r").Replace("\n", "\\n");
            return t.Length > 36 ? t.Substring(0, 33) + "..." : t;
        }

        private static string SafeFontString(Func<string> getter)
        {
            try
            {
                return getter() ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string GetEffectiveFontName(Word.Font font)
        {
            if (font == null)
            {
                return "";
            }

            string farEast = SafeFontString(() => font.NameFarEast);
            string name = SafeFontString(() => font.Name);
            if (!string.IsNullOrWhiteSpace(farEast))
            {
                return farEast.Trim();
            }

            return name;
        }

        private static void ApplyFontName(Word.Range range, string fontName)
        {
            if (string.IsNullOrEmpty(fontName))
            {
                return;
            }

            try
            {
                range.Font.NameFarEast = fontName;
            }
            catch (Exception ex)
            {
                DbgLog($"ApplyFontName NameFarEast 失败: {ex.Message}");
            }

            try
            {
                range.Font.Name = fontName;
            }
            catch (Exception ex)
            {
                DbgLog($"ApplyFontName Name 失败: {ex.Message}");
            }
        }

        private static void ApplyCharFormat(Word.Range range, Dictionary<string, object> format)
        {
            try
            {
                if (format.ContainsKey("font_name"))
                {
                    string fontName = format["font_name"]?.ToString() ?? "";
                    DbgLog($"ApplyCharFormat font_name=\"{fontName}\"");
                    ApplyFontName(range, fontName);
                }

                if (format.ContainsKey("font_size"))
                {
                    float fontSize = format["font_size"] is float f
                        ? f
                        : (format["font_size"] != null && float.TryParse(format["font_size"].ToString(), out float fs) ? fs : 0f);
                    if (fontSize > 0)
                    {
                        range.Font.Size = fontSize;
                    }
                }

                if (format.ContainsKey("font_color"))
                {
                    range.Font.Color = ConvertRgbToWdColor(format["font_color"]?.ToString() ?? "");
                }

                if (format.ContainsKey("background_color"))
                {
                    range.Shading.BackgroundPatternColor = ConvertRgbToWdColor(format["background_color"]?.ToString() ?? "");
                }

                if (format.ContainsKey("is_bold"))
                {
                    bool isBold = format["is_bold"] is bool b
                        ? b
                        : (format["is_bold"]?.ToString() == "True" || format["is_bold"]?.ToString() == "true");
                    range.Font.Bold = isBold ? 1 : 0;
                }

                if (format.ContainsKey("has_underline"))
                {
                    bool hasUnderline = format["has_underline"] is bool b
                        ? b
                        : (format["has_underline"]?.ToString() == "True" || format["has_underline"]?.ToString() == "true");
                    range.Font.Underline = hasUnderline
                        ? Word.WdUnderline.wdUnderlineSingle
                        : Word.WdUnderline.wdUnderlineNone;
                }

                if (format.ContainsKey("has_strikethrough"))
                {
                    bool hasStrikethrough = format["has_strikethrough"] is bool b
                        ? b
                        : (format["has_strikethrough"]?.ToString() == "True" || format["has_strikethrough"]?.ToString() == "true");
                    range.Font.StrikeThrough = hasStrikethrough ? 1 : 0;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"ApplyCharFormat failed: {ex.Message}", ex);
            }
        }

        private static string MapAlignment(Word.WdParagraphAlignment alignment)
        {
            switch (alignment)
            {
                case Word.WdParagraphAlignment.wdAlignParagraphCenter:
                    return "center";
                case Word.WdParagraphAlignment.wdAlignParagraphRight:
                    return "right";
                case Word.WdParagraphAlignment.wdAlignParagraphJustify:
                    return "justify";
                default:
                    return "left";
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
                default:
                    return Word.WdParagraphAlignment.wdAlignParagraphLeft;
            }
        }

        private static string ConvertWdColorToRgb(Word.WdColor wdColor)
        {
            try
            {
                if (wdColor == Word.WdColor.wdColorAutomatic || (int)wdColor == -16777216)
                {
                    return "automatic";
                }

                int colorValue = (int)wdColor;
                int b = (colorValue & 0xFF0000) >> 16;
                int g = (colorValue & 0x00FF00) >> 8;
                int r = colorValue & 0x0000FF;
                return $"#{r:X2}{g:X2}{b:X2}";
            }
            catch (Exception ex)
            {
                DbgLog($"ConvertWdColorToRgb: {ex.Message}");
                return "automatic";
            }
        }

        private static Word.WdColor ConvertRgbToWdColor(string rgbString)
        {
            try
            {
                if (string.IsNullOrEmpty(rgbString) || rgbString == "automatic")
                {
                    return Word.WdColor.wdColorAutomatic;
                }

                string hex = rgbString.TrimStart('#');
                if (hex.Length == 6)
                {
                    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                    int bgrValue = (b << 16) | (g << 8) | r;
                    return (Word.WdColor)bgrValue;
                }
            }
            catch (Exception ex)
            {
                DbgLog($"ConvertRgbToWdColor: {ex.Message}");
            }

            return Word.WdColor.wdColorAutomatic;
        }
    }
}
