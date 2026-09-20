using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    /// <summary>B4：段落格式方言（对齐/行距/段前段后/缩进/项目符号）。</summary>
    internal static class PptHtmlParagraphIo
    {
        public const string AlignLeft = "left";
        public const string AlignCenter = "center";
        public const string AlignRight = "right";
        public const string AlignJustify = "justify";

        public const string BulletNone = "none";
        public const string BulletDisc = "bullet";
        public const string BulletNumber = "number";

        // PowerPoint PpParagraphAlignment
        private const int PpAlignLeft = 1;
        private const int PpAlignCenter = 2;
        private const int PpAlignRight = 3;
        private const int PpAlignJustify = 4;
        private const int PpAlignDistribute = 5;

        // PowerPoint PpBulletType
        private const int PpBulletNone = 0;
        private const int PpBulletUnnumbered = 1;
        private const int PpBulletNumbered = 2;

        public sealed class Snapshot
        {
            public string Align { get; set; }

            /// <summary>倍数（如 1.15）或 exact:12（pt）。</summary>
            public string LineSpacing { get; set; }

            public double? SpaceBeforePt { get; set; }

            public double? SpaceAfterPt { get; set; }

            public double? IndentLeftPt { get; set; }

            public double? IndentFirstPt { get; set; }

            /// <summary>none / bullet / number</summary>
            public string Bullet { get; set; }
        }

        public static bool TryParseAlign(string raw, out string align, out string error)
        {
            align = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-align 为空";
                return false;
            }

            string s = raw.Trim().ToLowerInvariant();
            if (s == AlignLeft || s == AlignCenter || s == AlignRight || s == AlignJustify)
            {
                align = s;
                return true;
            }

            error = "非法 data-align（须为 left/center/right/justify）: " + raw;
            return false;
        }

        public static bool TryParseBullet(string raw, out string bullet, out string error)
        {
            bullet = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-bullet 为空";
                return false;
            }

            string s = raw.Trim().ToLowerInvariant();
            if (s == "true" || s == "disc" || s == BulletDisc)
            {
                bullet = BulletDisc;
                return true;
            }

            if (s == "false" || s == BulletNone)
            {
                bullet = BulletNone;
                return true;
            }

            if (s == BulletNumber || s == "numbered" || s == "decimal")
            {
                bullet = BulletNumber;
                return true;
            }

            error = "非法 data-bullet（须为 none/bullet/number）: " + raw;
            return false;
        }

        /// <summary>倍数如 1.5，或 exact:12（pt）。</summary>
        public static bool TryParseLineSpacing(string raw, out string normalized, out string error)
        {
            normalized = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-line-spacing 为空";
                return false;
            }

            string s = raw.Trim();
            if (s.StartsWith("exact:", StringComparison.OrdinalIgnoreCase))
            {
                string ptPart = s.Substring(6).Trim();
                if (ptPart.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
                {
                    ptPart = ptPart.Substring(0, ptPart.Length - 2).Trim();
                }

                if (!double.TryParse(ptPart, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt)
                    || pt <= 0
                    || pt > 400)
                {
                    error = "非法 data-line-spacing exact（须为 0–400 pt）: " + raw;
                    return false;
                }

                normalized = "exact:" + FormatPt(pt);
                return true;
            }

            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double mult)
                || mult <= 0
                || mult > 10)
            {
                error = "非法 data-line-spacing（须为 0–10 倍数，或 exact:N pt）: " + raw;
                return false;
            }

            normalized = FormatPt(mult);
            return true;
        }

        public static bool TryParseIndentOrSpacePt(string raw, string attrName, out double pt, out string error)
        {
            pt = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = attrName + " 为空";
                return false;
            }

            string s = raw.Trim();
            if (s.EndsWith("pt", StringComparison.OrdinalIgnoreCase))
            {
                s = s.Substring(0, s.Length - 2).Trim();
            }

            // 首行缩进可为负（悬挂）
            bool allowNeg = string.Equals(attrName, "data-indent-first", StringComparison.Ordinal);
            if (!double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out pt)
                || double.IsNaN(pt)
                || double.IsInfinity(pt)
                || pt > 400
                || pt < (allowNeg ? -400 : 0))
            {
                error = "非法 " + attrName + "（须为 pt 数值）: " + raw;
                return false;
            }

            return true;
        }

        public static string FormatPt(double pt)
        {
            return pt.ToString("0.##", CultureInfo.InvariantCulture);
        }

        public static string FormatAlignFromPp(int ppAlign)
        {
            switch (ppAlign)
            {
                case PpAlignCenter:
                    return AlignCenter;
                case PpAlignRight:
                    return AlignRight;
                case PpAlignJustify:
                case PpAlignDistribute:
                    return AlignJustify;
                default:
                    return AlignLeft;
            }
        }

        public static int AlignToPp(string align)
        {
            if (string.Equals(align, AlignCenter, StringComparison.OrdinalIgnoreCase))
            {
                return PpAlignCenter;
            }

            if (string.Equals(align, AlignRight, StringComparison.OrdinalIgnoreCase))
            {
                return PpAlignRight;
            }

            if (string.Equals(align, AlignJustify, StringComparison.OrdinalIgnoreCase))
            {
                return PpAlignJustify;
            }

            return PpAlignLeft;
        }

        public static Snapshot TryReadFromShape(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                // 缩进等仅在 TextFrame2/ParagraphFormat2；多段 Mixed 时读第 1 段
                Office.TextRange2 tr2 = shape.TextFrame2.TextRange;
                if (tr2 == null || tr2.Length < 1)
                {
                    return null;
                }

                Office.TextRange2 para = tr2.Paragraphs[1, 1];
                Office.ParagraphFormat2 pf = para.ParagraphFormat;
                var snap = new Snapshot();

                try
                {
                    snap.Align = FormatAlignFromPp((int)pf.Alignment);
                }
                catch (Exception)
                {
                }

                try
                {
                    float within = pf.SpaceWithin;
                    bool ruleLines = pf.LineRuleWithin == Office.MsoTriState.msoTrue;
                    if (ruleLines)
                    {
                        snap.LineSpacing = FormatPt(within);
                    }
                    else if (within > 0)
                    {
                        snap.LineSpacing = "exact:" + FormatPt(within);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    snap.SpaceBeforePt = pf.SpaceBefore;
                }
                catch (Exception)
                {
                }

                try
                {
                    snap.SpaceAfterPt = pf.SpaceAfter;
                }
                catch (Exception)
                {
                }

                try
                {
                    snap.IndentLeftPt = pf.LeftIndent;
                }
                catch (Exception)
                {
                }

                try
                {
                    snap.IndentFirstPt = pf.FirstLineIndent;
                }
                catch (Exception)
                {
                }

                try
                {
                    Office.BulletFormat2 bullet = pf.Bullet;
                    if (bullet.Visible != Office.MsoTriState.msoTrue)
                    {
                        snap.Bullet = BulletNone;
                    }
                    else
                    {
                        int bt = (int)bullet.Type;
                        if (bt == PpBulletNumbered)
                        {
                            snap.Bullet = BulletNumber;
                        }
                        else
                        {
                            snap.Bullet = BulletDisc;
                        }
                    }
                }
                catch (Exception)
                {
                    snap.Bullet = null;
                }

                return snap;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>现场段落已与 HTML 一致。未指定的字段不比。</summary>
        public static bool LiveMatches(
            Snapshot snap,
            string align,
            string lineSpacing,
            double? spaceBeforePt,
            double? spaceAfterPt,
            double? indentLeftPt,
            double? indentFirstPt,
            string bullet)
        {
            bool any = !string.IsNullOrEmpty(align)
                || !string.IsNullOrEmpty(lineSpacing)
                || spaceBeforePt.HasValue
                || spaceAfterPt.HasValue
                || indentLeftPt.HasValue
                || indentFirstPt.HasValue
                || !string.IsNullOrEmpty(bullet);
            if (!any)
            {
                return true;
            }

            if (snap == null)
            {
                return false;
            }

            if (!string.IsNullOrEmpty(align)
                && !string.Equals(snap.Align, align, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(lineSpacing)
                && !SameLineSpacing(snap.LineSpacing, lineSpacing))
            {
                return false;
            }

            if (spaceBeforePt.HasValue && !NearPt(snap.SpaceBeforePt, spaceBeforePt.Value))
            {
                return false;
            }

            if (spaceAfterPt.HasValue && !NearPt(snap.SpaceAfterPt, spaceAfterPt.Value))
            {
                return false;
            }

            if (indentLeftPt.HasValue && !NearPt(snap.IndentLeftPt, indentLeftPt.Value))
            {
                return false;
            }

            if (indentFirstPt.HasValue && !NearPt(snap.IndentFirstPt, indentFirstPt.Value))
            {
                return false;
            }

            if (!string.IsNullOrEmpty(bullet)
                && !string.Equals(snap.Bullet ?? BulletNone, bullet, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private static bool SameLineSpacing(string live, string want)
        {
            if (string.Equals(live, want, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (string.IsNullOrEmpty(live) || string.IsNullOrEmpty(want))
            {
                return false;
            }

            bool liveExact = live.StartsWith("exact:", StringComparison.OrdinalIgnoreCase);
            bool wantExact = want.StartsWith("exact:", StringComparison.OrdinalIgnoreCase);
            if (liveExact != wantExact)
            {
                return false;
            }

            string a = liveExact ? live.Substring(6) : live;
            string b = wantExact ? want.Substring(6) : want;
            if (!double.TryParse(a, NumberStyles.Float, CultureInfo.InvariantCulture, out double va)
                || !double.TryParse(b, NumberStyles.Float, CultureInfo.InvariantCulture, out double vb))
            {
                return false;
            }

            return Math.Abs(va - vb) <= 0.05;
        }

        private static bool NearPt(double? live, double want)
        {
            if (!live.HasValue)
            {
                return false;
            }

            return Math.Abs(live.Value - want) <= 0.2;
        }

        public static bool TryApplyToShape(
            PowerPoint.Shape shape,
            string align,
            string lineSpacing,
            double? spaceBeforePt,
            double? spaceAfterPt,
            double? indentLeftPt,
            double? indentFirstPt,
            string bullet,
            out string error)
        {
            error = null;
            if (shape == null)
            {
                return true;
            }

            bool any = !string.IsNullOrEmpty(align)
                || !string.IsNullOrEmpty(lineSpacing)
                || spaceBeforePt.HasValue
                || spaceAfterPt.HasValue
                || indentLeftPt.HasValue
                || indentFirstPt.HasValue
                || !string.IsNullOrEmpty(bullet);
            if (!any)
            {
                return true;
            }

            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return true;
                }

                // TextFrame2：对齐/行距/段距/缩进/项目符号（ParagraphFormat2）
                Office.ParagraphFormat2 pf = shape.TextFrame2.TextRange.ParagraphFormat;

                if (!string.IsNullOrEmpty(align))
                {
                    pf.Alignment = (Office.MsoParagraphAlignment)AlignToPp(align);
                }

                if (!string.IsNullOrEmpty(lineSpacing))
                {
                    if (lineSpacing.StartsWith("exact:", StringComparison.OrdinalIgnoreCase))
                    {
                        string ptPart = lineSpacing.Substring(6);
                        if (double.TryParse(ptPart, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt))
                        {
                            pf.LineRuleWithin = Office.MsoTriState.msoFalse;
                            pf.SpaceWithin = (float)pt;
                        }
                    }
                    else if (double.TryParse(lineSpacing, NumberStyles.Float, CultureInfo.InvariantCulture, out double mult))
                    {
                        pf.LineRuleWithin = Office.MsoTriState.msoTrue;
                        pf.SpaceWithin = (float)mult;
                    }
                }

                if (spaceBeforePt.HasValue)
                {
                    pf.SpaceBefore = (float)spaceBeforePt.Value;
                }

                if (spaceAfterPt.HasValue)
                {
                    pf.SpaceAfter = (float)spaceAfterPt.Value;
                }

                if (indentLeftPt.HasValue)
                {
                    pf.LeftIndent = (float)indentLeftPt.Value;
                }

                if (indentFirstPt.HasValue)
                {
                    pf.FirstLineIndent = (float)indentFirstPt.Value;
                }

                if (!string.IsNullOrEmpty(bullet))
                {
                    Office.BulletFormat2 bf = pf.Bullet;
                    if (string.Equals(bullet, BulletNone, StringComparison.OrdinalIgnoreCase))
                    {
                        bf.Visible = Office.MsoTriState.msoFalse;
                    }
                    else if (string.Equals(bullet, BulletNumber, StringComparison.OrdinalIgnoreCase))
                    {
                        bf.Visible = Office.MsoTriState.msoTrue;
                        bf.Type = Office.MsoBulletType.msoBulletNumbered;
                    }
                    else
                    {
                        bf.Visible = Office.MsoTriState.msoTrue;
                        bf.Type = Office.MsoBulletType.msoBulletUnnumbered;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写段落格式失败: " + ex.Message;
                return false;
            }
        }

        public static Snapshot TryReadFromWppShape(object shape)
        {
            if (shape == null)
            {
                return null;
            }

            try
            {
                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    return null;
                }

                // 优先 TextFrame2（含 LeftIndent/FirstLineIndent）
                object tr = null;
                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                if (tf2 != null)
                {
                    tr = WppCom.GetProperty(tf2, "TextRange");
                }

                if (tr == null)
                {
                    object tf = WppCom.GetProperty(shape, "TextFrame");
                    tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                }

                if (tr == null)
                {
                    return null;
                }

                object lenObj = WppCom.GetProperty(tr, "Length");
                int len = lenObj == null ? 0 : Convert.ToInt32(lenObj);
                if (len < 1)
                {
                    return null;
                }

                object para = null;
                try
                {
                    para = WppCom.Invoke(tr, "Paragraphs", 1, 1);
                }
                catch (Exception)
                {
                    try
                    {
                        para = WppCom.Invoke(tr, "Paragraphs", 1);
                    }
                    catch (Exception)
                    {
                        para = tr;
                    }
                }

                object pf = para == null ? null : WppCom.GetProperty(para, "ParagraphFormat");
                if (pf == null)
                {
                    return null;
                }

                var snap = new Snapshot();
                try
                {
                    object al = WppCom.GetProperty(pf, "Alignment");
                    if (al != null)
                    {
                        snap.Align = FormatAlignFromPp(Convert.ToInt32(al));
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    // WPP：Rule=-1/1 表示倍数（读侧）；写侧勿按 Office 字面去写规则，倍数只写 SpaceWithin。
                    object within = WppCom.GetProperty(pf, "SpaceWithin");
                    object rule = WppCom.GetProperty(pf, "LineRuleWithin");
                    float w = within == null ? 0f : Convert.ToSingle(within);
                    int ruleInt = rule == null ? int.MinValue : Convert.ToInt32(rule);
                    bool ruleLines = rule != null && (ruleInt == -1 || ruleInt == 1);
                    if (ruleLines)
                    {
                        snap.LineSpacing = FormatPt(w);
                    }
                    else if (w > 0)
                    {
                        snap.LineSpacing = "exact:" + FormatPt(w);
                    }

                    if (DebugLineSpacing)
                    {
                        Console.WriteLine(
                            "[LS-DEBUG] WPP读 ParagraphFormat"
                            + " LineRuleWithin=" + (rule == null ? "null" : ruleInt.ToString(CultureInfo.InvariantCulture))
                            + " SpaceWithin=" + w.ToString(CultureInfo.InvariantCulture)
                            + " → dialect=" + (snap.LineSpacing ?? "(空)"));
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object v = WppCom.GetProperty(pf, "SpaceBefore");
                    if (v != null)
                    {
                        snap.SpaceBeforePt = Convert.ToDouble(v);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object v = WppCom.GetProperty(pf, "SpaceAfter");
                    if (v != null)
                    {
                        snap.SpaceAfterPt = Convert.ToDouble(v);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object v = WppCom.GetProperty(pf, "LeftIndent");
                    if (v != null)
                    {
                        snap.IndentLeftPt = Convert.ToDouble(v);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object v = WppCom.GetProperty(pf, "FirstLineIndent");
                    if (v != null)
                    {
                        snap.IndentFirstPt = Convert.ToDouble(v);
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    object bullet = WppCom.GetProperty(pf, "Bullet");
                    object vis = bullet == null ? null : WppCom.GetProperty(bullet, "Visible");
                    bool on = vis != null && (Convert.ToInt32(vis) == -1 || Convert.ToInt32(vis) == 1);
                    if (!on)
                    {
                        snap.Bullet = BulletNone;
                    }
                    else
                    {
                        object bt = WppCom.GetProperty(bullet, "Type");
                        int t = bt == null ? PpBulletUnnumbered : Convert.ToInt32(bt);
                        snap.Bullet = t == PpBulletNumbered ? BulletNumber : BulletDisc;
                    }
                }
                catch (Exception)
                {
                }

                return snap;
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool TryApplyToWppShape(
            object shape,
            string align,
            string lineSpacing,
            double? spaceBeforePt,
            double? spaceAfterPt,
            double? indentLeftPt,
            double? indentFirstPt,
            string bullet,
            out string error)
        {
            error = null;
            if (shape == null)
            {
                return true;
            }

            bool any = !string.IsNullOrEmpty(align)
                || !string.IsNullOrEmpty(lineSpacing)
                || spaceBeforePt.HasValue
                || spaceAfterPt.HasValue
                || indentLeftPt.HasValue
                || indentFirstPt.HasValue
                || !string.IsNullOrEmpty(bullet);
            if (!any)
            {
                return true;
            }

            try
            {
                if (!TryGetWppTextRange(shape, out object tr) || tr == null)
                {
                    return true;
                }

                // 与读路径一致：优先逐段 ParagraphFormat；整段 TextRange.ParagraphFormat 在 WPS 上
                // 常写不住 LineRuleWithin（倍数会落成定距 pt）。
                var pfs = new List<object>();
                CollectWppParagraphFormats(shape, tr, pfs);
                if (pfs.Count == 0)
                {
                    return true;
                }

                if (DebugLineSpacing)
                {
                    Console.WriteLine("[LS-DEBUG] WPP写命中 ParagraphFormat ×" + pfs.Count);
                }

                bool multipleOk = false;
                if (!string.IsNullOrEmpty(lineSpacing)
                    && !lineSpacing.StartsWith("exact:", StringComparison.OrdinalIgnoreCase)
                    && double.TryParse(
                        lineSpacing,
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out double multWant))
                {
                    // 先只写 SpaceWithin；失败则多目标回退（仍只写 SpaceWithin，见 TryApplyWppMultipleFallback）
                    multipleOk = TryApplyWppMultipleLineSpacing(shape, multWant)
                        || TryApplyWppMultipleFallback(shape, multWant);
                }

                for (int i = 0; i < pfs.Count; i++)
                {
                    object pf = pfs[i];
                    if (pf == null)
                    {
                        continue;
                    }

                    if (!string.IsNullOrEmpty(align))
                    {
                        TrySet(pf, "Alignment", AlignToPp(align));
                    }

                    if (!string.IsNullOrEmpty(lineSpacing) && !multipleOk)
                    {
                        ApplyWppLineSpacing(pf, lineSpacing);
                    }

                    if (spaceBeforePt.HasValue)
                    {
                        TrySet(pf, "SpaceBefore", spaceBeforePt.Value);
                    }

                    if (spaceAfterPt.HasValue)
                    {
                        TrySet(pf, "SpaceAfter", spaceAfterPt.Value);
                    }

                    if (indentLeftPt.HasValue)
                    {
                        TrySet(pf, "LeftIndent", indentLeftPt.Value);
                    }

                    if (indentFirstPt.HasValue)
                    {
                        TrySet(pf, "FirstLineIndent", indentFirstPt.Value);
                    }

                    if (!string.IsNullOrEmpty(bullet))
                    {
                        object bf = WppCom.GetProperty(pf, "Bullet");
                        if (string.Equals(bullet, BulletNone, StringComparison.OrdinalIgnoreCase))
                        {
                            TrySet(bf, "Visible", 0);
                        }
                        else if (string.Equals(bullet, BulletNumber, StringComparison.OrdinalIgnoreCase))
                        {
                            TrySet(bf, "Visible", -1);
                            TrySet(bf, "Type", PpBulletNumbered);
                        }
                        else
                        {
                            TrySet(bf, "Visible", -1);
                            TrySet(bf, "Type", PpBulletUnnumbered);
                        }
                    }

                    if (!string.IsNullOrEmpty(lineSpacing))
                    {
                        DumpWppLineSpacing(
                            pf,
                            multipleOk ? "WPP写段落结束#策略成功#" + (i + 1) : "WPP写段落结束#" + (i + 1),
                            lineSpacing);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "写段落格式失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryGetWppTextRange(object shape, out object textRange)
        {
            textRange = null;
            if (shape == null || !IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
            {
                return false;
            }

            object tf2 = WppCom.GetProperty(shape, "TextFrame2");
            if (tf2 != null)
            {
                textRange = WppCom.GetProperty(tf2, "TextRange");
            }

            if (textRange == null)
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                textRange = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
            }

            return textRange != null;
        }

        /// <summary>收集 WPS 上可能生效的 ParagraphFormat：TextFrame2/TextFrame 整段 + 各段。</summary>
        private static void CollectWppParagraphFormats(object shape, object textRange, List<object> into)
        {
            if (into == null)
            {
                return;
            }

            // 1) 调用方已解析的 TextRange（通常 TextFrame2）
            CollectWppParagraphFormatsFromTextRange(textRange, into);

            // 2) 经典 TextFrame.TextRange —— 部分 WPS 版本只在这里认 LineRuleWithin
            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object trLegacy = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                if (trLegacy != null && !ReferenceEquals(trLegacy, textRange))
                {
                    CollectWppParagraphFormatsFromTextRange(trLegacy, into);
                }
            }
            catch (Exception)
            {
            }
        }

        private static void CollectWppParagraphFormatsFromTextRange(object textRange, List<object> into)
        {
            if (textRange == null || into == null)
            {
                return;
            }

            // 整段 ParagraphFormat
            try
            {
                object pfAll = WppCom.GetProperty(textRange, "ParagraphFormat");
                AddUniqueCom(into, pfAll);
            }
            catch (Exception)
            {
            }

            for (int i = 1; i <= 64; i++)
            {
                object para = null;
                try
                {
                    para = WppCom.Invoke(textRange, "Paragraphs", i, 1);
                }
                catch (Exception)
                {
                    try
                    {
                        para = WppCom.Invoke(textRange, "Paragraphs", i);
                    }
                    catch (Exception)
                    {
                        break;
                    }
                }

                if (para == null)
                {
                    break;
                }

                try
                {
                    object pf = WppCom.GetProperty(para, "ParagraphFormat");
                    AddUniqueCom(into, pf);
                }
                catch (Exception)
                {
                }
            }
        }

        private static void AddUniqueCom(List<object> into, object item)
        {
            if (item == null || into == null)
            {
                return;
            }

            for (int i = 0; i < into.Count; i++)
            {
                if (ReferenceEquals(into[i], item))
                {
                    return;
                }
            }

            into.Add(item);
        }

        /// <summary>
        /// WPP 阶梯探针结论（晚绑定）：
        /// - 只写 SpaceWithin → 读回多为倍数（Rule=-1）
        /// - 写 LineRuleWithin=-1/1（本意倍数）→ 读回反而定距（Rule=0）
        /// - 写 LineRuleWithin=0（本意定距）→ 常钉不住，仍是倍数
        /// 故倍数路径：只写 SpaceWithin，绝不碰 LineRuleWithin。
        /// </summary>
        private static bool TryApplyWppMultipleLineSpacing(object shape, double mult)
        {
            if (shape == null || mult <= 0)
            {
                return false;
            }

            float m = (float)mult;
            string tag = mult.ToString("0.##", CultureInfo.InvariantCulture);
            object pf = TryWppPf_TextFrameWhole(shape, "quick");
            if (pf == null)
            {
                pf = TryWppPf_TextFrame2Whole(shape, "quick2");
            }

            if (pf == null)
            {
                return false;
            }

            // 关键：不要写 LineRuleWithin（写 -1/1 会把模式拧成定距）
            TrySetLogged(pf, "SpaceWithin", m);
            DumpWppLineSpacing(pf, "WPP只写SpaceWithin(倍数)", tag);

            if (IsWppMultipleLive(shape, mult, out string how))
            {
                if (DebugLineSpacing)
                {
                    Console.WriteLine("[LS-DEBUG] WPP真倍数成功 want=" + tag + " via " + how);
                }

                return true;
            }

            if (DebugLineSpacing)
            {
                DumpWppLineSpacing(pf, "WPP只写SpaceWithin未活，将多目标回退", tag);
            }

            return false;
        }

        /// <summary>
        /// 多 ParagraphFormat 目标只写 SpaceWithin（仍不碰 LineRuleWithin）。
        /// 若已在倍数态，可先写大诱饵再改回真倍数，避免某路径仍停在定距小数。
        /// </summary>
        private static bool TryApplyWppMultipleFallback(object shape, double mult)
        {
            if (shape == null || mult <= 0)
            {
                return false;
            }

            double fontPt = TryReadWppFontSizePt(shape);
            if (fontPt < 1.0)
            {
                fontPt = 18.0;
            }

            double baitPt = fontPt * mult;
            if (baitPt < 1.0)
            {
                baitPt = 1.0;
            }

            float multF = (float)mult;
            string want = mult.ToString("0.##", CultureInfo.InvariantCulture);

            if (DebugLineSpacing)
            {
                Console.WriteLine(
                    "[LS-DEBUG] WPP倍数多目标回退 mult="
                    + want
                    + " fontPt=" + fontPt.ToString("0.##", CultureInfo.InvariantCulture)
                    + " baitPt=" + baitPt.ToString("0.##", CultureInfo.InvariantCulture));
            }

            object[] targets =
            {
                TryWppPf_TextFrameWhole(shape, "fb-tf"),
                TryWppPf_TextFrame2Whole(shape, "fb-tf2"),
                TryWppPf_TextFramePara1(shape, "fb-p1"),
                TryWppPf_TextFrame2Para1(shape, "fb-p2")
            };

            bool any = false;
            for (int i = 0; i < targets.Length; i++)
            {
                object pf = targets[i];
                if (pf == null)
                {
                    continue;
                }

                // 先只写真倍数
                TrySetLogged(pf, "SpaceWithin", multF);
                DumpWppLineSpacing(pf, "WPP回退直写#" + (i + 1), want);

                if (TryPeekWppLineRule(pf, out int rule, out float within)
                    && IsLineRuleMultiple(rule)
                    && Math.Abs(within - multF) <= 0.08f)
                {
                    any = true;
                    continue;
                }

                // 诱饵拉到倍数态后，只改 SpaceWithin（禁止再写 LineRuleWithin）
                TrySetLogged(pf, "SpaceWithin", (float)baitPt);
                DumpWppLineSpacing(pf, "WPP回退诱饵#" + (i + 1), want);
                TrySetLogged(pf, "SpaceWithin", multF);
                DumpWppLineSpacing(pf, "WPP回退改回#" + (i + 1), want);
                any = true;
            }

            return any;
        }

        private static double TryReadWppFontSizePt(object shape)
        {
            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object size = font == null ? null : WppCom.GetProperty(font, "Size");
                if (size != null)
                {
                    double pt = Convert.ToDouble(size);
                    if (pt >= 1.0)
                    {
                        return pt;
                    }
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                object tr = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                object font = tr == null ? null : WppCom.GetProperty(tr, "Font");
                object size = font == null ? null : WppCom.GetProperty(font, "Size");
                if (size != null)
                {
                    double pt = Convert.ToDouble(size);
                    if (pt >= 1.0)
                    {
                        return pt;
                    }
                }
            }
            catch (Exception)
            {
            }

            return 0;
        }

        private static bool IsWppMultipleLive(object shape, double mult, out string how)
        {
            how = null;
            try
            {
                Snapshot snap = TryReadFromWppShape(shape);
                if (snap == null || string.IsNullOrEmpty(snap.LineSpacing))
                {
                    return false;
                }

                string ls = snap.LineSpacing.Trim();
                if (ls.StartsWith("exact:", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                if (!double.TryParse(ls, NumberStyles.Float, CultureInfo.InvariantCulture, out double got))
                {
                    return false;
                }

                if (Math.Abs(got - mult) <= 0.08)
                {
                    how = "read dialect=" + ls;
                    return true;
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static object TryWppPf_TextFrameWhole(object shape, string name)
        {
            if (DebugLineSpacing)
            {
                Console.WriteLine("[LS-DEBUG] " + name + " TextFrame.TextRange.ParagraphFormat");
            }

            object tf = WppCom.GetProperty(shape, "TextFrame");
            object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
            return tr == null ? null : WppCom.GetProperty(tr, "ParagraphFormat");
        }

        private static object TryWppPf_TextFrame2Whole(object shape, string name)
        {
            if (DebugLineSpacing)
            {
                Console.WriteLine("[LS-DEBUG] " + name + " TextFrame2.TextRange.ParagraphFormat");
            }

            object tf2 = WppCom.GetProperty(shape, "TextFrame2");
            object tr = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
            return tr == null ? null : WppCom.GetProperty(tr, "ParagraphFormat");
        }

        private static object TryWppPf_TextFramePara1(object shape, string name)
        {
            if (DebugLineSpacing)
            {
                Console.WriteLine("[LS-DEBUG] " + name + " TextFrame.Paragraphs(1)");
            }

            object tf = WppCom.GetProperty(shape, "TextFrame");
            object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
            return TryWppPf_FromParagraph1(tr);
        }

        private static object TryWppPf_TextFrame2Para1(object shape, string name)
        {
            if (DebugLineSpacing)
            {
                Console.WriteLine("[LS-DEBUG] " + name + " TextFrame2.Paragraphs(1)");
            }

            object tf2 = WppCom.GetProperty(shape, "TextFrame2");
            object tr = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
            return TryWppPf_FromParagraph1(tr);
        }

        private static object TryWppPf_FromParagraph1(object tr)
        {
            if (tr == null)
            {
                return null;
            }

            object para = null;
            try
            {
                para = WppCom.Invoke(tr, "Paragraphs", 1, 1);
            }
            catch (Exception)
            {
                try
                {
                    para = WppCom.Invoke(tr, "Paragraphs", 1);
                }
                catch (Exception)
                {
                }
            }

            return para == null ? null : WppCom.GetProperty(para, "ParagraphFormat");
        }

        private static object TryWppPf_SelectThenTextFrame(object shape, string name)
        {
            try
            {
                WppCom.Invoke(shape, "Select");
            }
            catch (Exception)
            {
            }

            object tf = WppCom.GetProperty(shape, "TextFrame");
            object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
            try
            {
                if (tr != null)
                {
                    WppCom.Invoke(tr, "Select");
                }
            }
            catch (Exception)
            {
            }

            return tr == null ? null : WppCom.GetProperty(tr, "ParagraphFormat");
        }

        private static object TryWppPf_GrowBoxThenTextFrame(object shape, string name)
        {
            try
            {
                object hObj = WppCom.GetProperty(shape, "Height");
                double oldH = hObj == null ? 0 : Convert.ToDouble(hObj);
                if (oldH > 0)
                {
                    WppCom.TrySetProperty(shape, "Height", oldH * 2.5);
                }
            }
            catch (Exception)
            {
            }

            return TryWppPf_TextFrameWhole(shape, name + "-grow");
        }

        /// <summary>
        /// WPP 行距写入（阶梯探针）：
        /// 倍数 → 只写 SpaceWithin；定距 → 先写 LineRuleWithin=-1/1（会把读回拧成定距），再写 pt。
        /// </summary>
        private static void ApplyWppLineSpacing(object pf, string lineSpacing)
        {
            if (pf == null || string.IsNullOrEmpty(lineSpacing))
            {
                return;
            }

            if (lineSpacing.StartsWith("exact:", StringComparison.OrdinalIgnoreCase))
            {
                string ptPart = lineSpacing.Substring(6);
                if (!double.TryParse(ptPart, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt))
                {
                    return;
                }

                // 探针：写 -1/1 反而落成定距；写 0 钉不住
                TryForceWppLineRuleExactViaInvertedWrite(pf);
                TrySet(pf, "SpaceWithin", (float)pt);
                TryForceWppLineRuleExactViaInvertedWrite(pf);
                DumpWppLineSpacing(pf, "WPP写定距", lineSpacing);
                return;
            }

            if (!double.TryParse(lineSpacing, NumberStyles.Float, CultureInfo.InvariantCulture, out double mult))
            {
                return;
            }

            TrySetLogged(pf, "SpaceWithin", (float)mult);
            DumpWppLineSpacing(pf, "WPP写倍数(只SpaceWithin)", lineSpacing);
        }

        /// <summary>
        /// WPP 上要把模式拧成定距，应写 LineRuleWithin=-1/1（读回变为 0）；
        /// 写 0 往往钉不住。与 Office 字面语义相反。
        /// </summary>
        private static bool TryForceWppLineRuleExactViaInvertedWrite(object pf)
        {
            return TryForceWppLineRuleMultiple(pf);
        }

        private static void DumpWppParagraphFormatMembers(object pf)
        {
            if (!DebugLineSpacing || pf == null)
            {
                return;
            }

            try
            {
                Console.WriteLine("[LS-DEBUG] ParagraphFormat type=" + pf.GetType().FullName);
                foreach (string name in new[]
                {
                    "LineRuleWithin", "SpaceWithin", "LineSpacingRule", "LineSpacing",
                    "SpaceBefore", "SpaceAfter"
                })
                {
                    try
                    {
                        object v = WppCom.GetProperty(pf, name);
                        Console.WriteLine(
                            "[LS-DEBUG]   get " + name + "="
                            + (v == null ? "null" : v.ToString() + " (" + v.GetType().Name + ")"));
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(
                            "[LS-DEBUG]   get " + name + " ERR "
                            + ex.GetBaseException().Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("[LS-DEBUG] dump members 失败: " + ex.Message);
            }
        }

        private static bool TryForceWppLineRuleMultiple(object pf)
        {
            // msoTrue 在晚绑定里常见 -1；个别宿主只认 1 / true / short
            // 另：部分 WPS 对 SetProperty 静默吞掉，需换 put_ 名或 Invoke
            object[] candidates = { -1, 1, true, (short)-1, (short)1 };
            string[] names = { "LineRuleWithin", "put_LineRuleWithin" };
            for (int n = 0; n < names.Length; n++)
            {
                for (int i = 0; i < candidates.Length; i++)
                {
                    if (!TrySetLogged(pf, names[n], candidates[i]))
                    {
                        continue;
                    }

                    if (TryPeekWppLineRule(pf, out int rule, out _) && IsLineRuleMultiple(rule))
                    {
                        return true;
                    }

                    if (DebugLineSpacing)
                    {
                        Console.WriteLine(
                            "[LS-DEBUG] 设 " + names[n] + "=" + candidates[i]
                            + " 后仍非倍数 rule="
                            + (TryPeekWppLineRule(pf, out int r2, out _)
                                ? r2.ToString(CultureInfo.InvariantCulture)
                                : "?"));
                    }
                }
            }

            return false;
        }

        private static bool TrySetLogged(object target, string name, object value)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return false;
            }

            try
            {
                target.GetType().InvokeMember(
                    name,
                    System.Reflection.BindingFlags.SetProperty
                        | System.Reflection.BindingFlags.Instance
                        | System.Reflection.BindingFlags.Public,
                    null,
                    target,
                    new object[] { value });
                return true;
            }
            catch (Exception ex)
            {
                if (DebugLineSpacing)
                {
                    Console.WriteLine(
                        "[LS-DEBUG] Set " + name + "=" + value
                        + " 失败: " + ex.GetBaseException().Message);
                }

                return false;
            }
        }

        /// <summary>环境变量 PPT_HTML_DEBUG_LINESPACING=1 时打印 WPP 行距 COM 读写探针。</summary>
        internal static bool DebugLineSpacing
        {
            get
            {
                string v = Environment.GetEnvironmentVariable("PPT_HTML_DEBUG_LINESPACING");
                return string.Equals(v, "1", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(v, "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool IsLineRuleMultiple(int rule)
        {
            return rule == -1 || rule == 1;
        }

        private static bool TryPeekWppLineRule(object pf, out int rule, out float within)
        {
            rule = int.MinValue;
            within = 0f;
            try
            {
                object r = WppCom.GetProperty(pf, "LineRuleWithin");
                object w = WppCom.GetProperty(pf, "SpaceWithin");
                if (r != null)
                {
                    rule = Convert.ToInt32(r);
                }

                if (w != null)
                {
                    within = Convert.ToSingle(w);
                }

                return r != null || w != null;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void DumpWppLineSpacing(object pf, string stage, string want)
        {
            if (!DebugLineSpacing || pf == null)
            {
                return;
            }

            try
            {
                if (!TryPeekWppLineRule(pf, out int rule, out float within))
                {
                    Console.WriteLine("[LS-DEBUG] " + stage + " want=" + want + " （读 COM 失败）");
                    return;
                }

                string kind = IsLineRuleMultiple(rule) ? "倍数" : "定距(exact)";
                string asDialect = IsLineRuleMultiple(rule)
                    ? FormatPt(within)
                    : (within > 0 ? "exact:" + FormatPt(within) : "(空)");
                Console.WriteLine(
                    "[LS-DEBUG] " + stage
                    + " want=" + want
                    + " LineRuleWithin=" + rule
                    + " SpaceWithin=" + within.ToString(CultureInfo.InvariantCulture)
                    + " →" + kind
                    + " dialect=" + asDialect);
            }
            catch (Exception ex)
            {
                Console.WriteLine("[LS-DEBUG] " + stage + " dump失败: " + ex.Message);
            }
        }

        private static bool IsTruthy(object value)
        {
            if (value == null)
            {
                return false;
            }

            try
            {
                int n = Convert.ToInt32(value);
                return n == -1 || n == 1;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void TrySet(object target, string name, object value)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(target, name, value);
            }
            catch (Exception)
            {
            }
        }
    }
}
