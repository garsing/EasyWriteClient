using System;
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
                    object within = WppCom.GetProperty(pf, "SpaceWithin");
                    object rule = WppCom.GetProperty(pf, "LineRuleWithin");
                    float w = within == null ? 0f : Convert.ToSingle(within);
                    bool ruleLines = rule != null && (Convert.ToInt32(rule) == -1 || Convert.ToInt32(rule) == 1);
                    if (ruleLines)
                    {
                        snap.LineSpacing = FormatPt(w);
                    }
                    else if (w > 0)
                    {
                        snap.LineSpacing = "exact:" + FormatPt(w);
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
                if (!IsTruthy(WppCom.GetProperty(shape, "HasTextFrame")))
                {
                    return true;
                }

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

                object pf = tr == null ? null : WppCom.GetProperty(tr, "ParagraphFormat");
                if (pf == null)
                {
                    return true;
                }

                if (!string.IsNullOrEmpty(align))
                {
                    TrySet(pf, "Alignment", AlignToPp(align));
                }

                if (!string.IsNullOrEmpty(lineSpacing))
                {
                    if (lineSpacing.StartsWith("exact:", StringComparison.OrdinalIgnoreCase))
                    {
                        string ptPart = lineSpacing.Substring(6);
                        if (double.TryParse(ptPart, NumberStyles.Float, CultureInfo.InvariantCulture, out double pt))
                        {
                            TrySet(pf, "LineRuleWithin", 0);
                            TrySet(pf, "SpaceWithin", pt);
                        }
                    }
                    else if (double.TryParse(lineSpacing, NumberStyles.Float, CultureInfo.InvariantCulture, out double mult))
                    {
                        TrySet(pf, "LineRuleWithin", -1);
                        TrySet(pf, "SpaceWithin", mult);
                    }
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

                return true;
            }
            catch (Exception ex)
            {
                error = "写段落格式失败: " + ex.Message;
                return false;
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
