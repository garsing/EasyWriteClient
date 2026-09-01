using System;
using System.Globalization;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlTextWidthIo
    {
        public static string FormatPct(double pct)
        {
            return pct.ToString("0.0", CultureInfo.InvariantCulture);
        }

        public static bool TryMeasurePowerPoint(
            PowerPoint.Shape shape,
            float slideWidth,
            out double widthPct)
        {
            widthPct = 0;
            if (shape == null || slideWidth <= 0)
            {
                return false;
            }

            try
            {
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                {
                    return false;
                }

                string text = shape.TextFrame.TextRange.Text;
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                float bound = shape.TextFrame2.TextRange.BoundWidth;
                if (bound <= 0)
                {
                    return false;
                }

                widthPct = bound / slideWidth * 100.0;
                if (widthPct < 0)
                {
                    widthPct = 0;
                }

                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public static bool TryMeasureWpp(object shape, float slideWidth, out double widthPct)
        {
            widthPct = 0;
            if (shape == null || slideWidth <= 0)
            {
                return false;
            }

            try
            {
                object tf = WppCom.GetProperty(shape, "TextFrame");
                object tr = tf == null ? null : WppCom.GetProperty(tf, "TextRange");
                string text = tr == null ? null : Convert.ToString(WppCom.GetProperty(tr, "Text"));
                if (string.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                object tf2 = WppCom.GetProperty(shape, "TextFrame2");
                object tr2 = tf2 == null ? null : WppCom.GetProperty(tf2, "TextRange");
                object boundObj = tr2 == null ? null : WppCom.GetProperty(tr2, "BoundWidth");
                if (boundObj == null)
                {
                    return false;
                }

                float bound = Convert.ToSingle(boundObj, CultureInfo.InvariantCulture);
                if (bound <= 0)
                {
                    return false;
                }

                widthPct = bound / slideWidth * 100.0;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}
