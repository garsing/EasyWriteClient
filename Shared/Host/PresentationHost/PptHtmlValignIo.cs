using System;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlValignIo
    {
        public const string Top = "top";
        public const string Middle = "middle";
        public const string Bottom = "bottom";

        private const int MsoAnchorTop = 1;
        private const int MsoAnchorMiddle = 3;
        private const int MsoAnchorBottom = 4;

        public static bool TryParse(string raw, out string valign, out string error)
        {
            valign = null;
            error = null;
            if (string.IsNullOrWhiteSpace(raw))
            {
                error = "data-valign 为空";
                return false;
            }

            string s = raw.Trim().ToLowerInvariant();
            if (s == Top || s == Middle || s == Bottom)
            {
                valign = s;
                return true;
            }

            error = "非法 data-valign（须为 top/middle/bottom）: " + raw;
            return false;
        }

        public static string FromAnchor(int anchor)
        {
            if (anchor == MsoAnchorMiddle)
            {
                return Middle;
            }

            if (anchor == MsoAnchorBottom)
            {
                return Bottom;
            }

            if (anchor == MsoAnchorTop)
            {
                return Top;
            }

            return null;
        }

        public static int ToAnchor(string valign)
        {
            if (valign == Middle)
            {
                return MsoAnchorMiddle;
            }

            if (valign == Bottom)
            {
                return MsoAnchorBottom;
            }

            return MsoAnchorTop;
        }

        public static string TryReadPowerPoint(PowerPoint.Shape shape)
        {
            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return null;
                }

                return FromAnchor((int)shape.TextFrame2.VerticalAnchor);
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool TryWritePowerPoint(PowerPoint.Shape shape, string valign, out string warning)
        {
            warning = null;
            try
            {
                shape.TextFrame2.VerticalAnchor = (Office.MsoVerticalAnchor)ToAnchor(valign);
                return true;
            }
            catch (Exception ex)
            {
                warning = "无法写入 data-valign=" + valign + ": " + ex.Message;
                return false;
            }
        }

        public static string TryReadWpp(object shape)
        {
            try
            {
                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                object raw = tf2 == null ? null : WppCom.GetProperty(tf2, "VerticalAnchor");
                if (raw == null)
                {
                    return null;
                }

                return FromAnchor(Convert.ToInt32(raw));
            }
            catch (Exception)
            {
                return null;
            }
        }

        public static bool TryWriteWpp(object shape, string valign, out string warning)
        {
            warning = null;
            try
            {
                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                if (tf2 == null)
                {
                    warning = "wpp 订不上 TextFrame2.VerticalAnchor";
                    return false;
                }

                WppCom.TrySetProperty(tf2, "VerticalAnchor", ToAnchor(valign));
                string back = TryReadWpp(shape);
                if (!string.Equals(back, valign, StringComparison.OrdinalIgnoreCase))
                {
                    warning = "wpp 订不上 data-valign=" + valign;
                    return false;
                }

                return true;
            }
            catch (Exception)
            {
                warning = "wpp 订不上 data-valign=" + valign;
                return false;
            }
        }
    }
}
