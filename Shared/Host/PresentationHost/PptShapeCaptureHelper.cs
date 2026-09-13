using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 整页导出图按形状外接矩形裁切，并按 D17 做极少 padding / 过小上采样。
    /// </summary>
    internal static class PptShapeCaptureHelper
    {
        public sealed class CropResult
        {
            public Bitmap Bitmap { get; set; }
            public double Scale { get; set; }
            public double LeftPct { get; set; }
            public double TopPct { get; set; }
            public double WidthPct { get; set; }
            public double HeightPct { get; set; }
        }

        public static bool TryCropFromSlidePng(
            string pngPath,
            float slideWidthPt,
            float slideHeightPt,
            float shapeLeftPt,
            float shapeTopPt,
            float shapeWidthPt,
            float shapeHeightPt,
            out CropResult crop,
            out string error)
        {
            crop = null;
            error = null;
            if (string.IsNullOrWhiteSpace(pngPath) || !File.Exists(pngPath))
            {
                error = "幻灯片导出图片不存在";
                return false;
            }

            Bitmap page = null;
            try
            {
                page = new Bitmap(pngPath);
                return TryCropFromSlideBitmap(
                    page,
                    slideWidthPt,
                    slideHeightPt,
                    shapeLeftPt,
                    shapeTopPt,
                    shapeWidthPt,
                    shapeHeightPt,
                    out crop,
                    out error);
            }
            catch (Exception ex)
            {
                error = "裁切形状失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (page != null)
                {
                    page.Dispose();
                }
            }
        }

        /// <summary>
        /// 已加载的整页位图上裁切形状（同页多 shape 时只 Export 一次）。
        /// </summary>
        public static bool TryCropFromSlideBitmap(
            Bitmap page,
            float slideWidthPt,
            float slideHeightPt,
            float shapeLeftPt,
            float shapeTopPt,
            float shapeWidthPt,
            float shapeHeightPt,
            out CropResult crop,
            out string error)
        {
            crop = null;
            error = null;
            if (page == null)
            {
                error = "幻灯片导出图片不存在";
                return false;
            }

            if (slideWidthPt <= 0f || slideHeightPt <= 0f)
            {
                error = "幻灯片页面尺寸无效";
                return false;
            }

            if (shapeWidthPt <= 0f || shapeHeightPt <= 0f)
            {
                error = "形状尺寸无效";
                return false;
            }

            try
            {
                int pageW = page.Width;
                int pageH = page.Height;
                if (pageW < 1 || pageH < 1)
                {
                    error = "幻灯片导出图片为空";
                    return false;
                }

                double leftPx = shapeLeftPt / slideWidthPt * pageW;
                double topPx = shapeTopPt / slideHeightPt * pageH;
                double widthPx = shapeWidthPt / slideWidthPt * pageW;
                double heightPx = shapeHeightPt / slideHeightPt * pageH;

                double pad = Math.Max(2.0, Math.Round(0.003 * pageW));
                leftPx -= pad;
                topPx -= pad;
                widthPx += pad * 2;
                heightPx += pad * 2;

                int x0 = (int)Math.Floor(leftPx);
                int y0 = (int)Math.Floor(topPx);
                int x1 = (int)Math.Ceiling(leftPx + widthPx);
                int y1 = (int)Math.Ceiling(topPx + heightPx);
                x0 = Math.Max(0, Math.Min(x0, pageW - 1));
                y0 = Math.Max(0, Math.Min(y0, pageH - 1));
                x1 = Math.Max(x0 + 1, Math.Min(x1, pageW));
                y1 = Math.Max(y0 + 1, Math.Min(y1, pageH));
                int cropW = x1 - x0;
                int cropH = y1 - y0;
                if (cropW < 1 || cropH < 1)
                {
                    error = "形状裁切区域无效";
                    return false;
                }

                var cropped = new Bitmap(cropW, cropH, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(cropped))
                {
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.DrawImage(
                        page,
                        new Rectangle(0, 0, cropW, cropH),
                        new Rectangle(x0, y0, cropW, cropH),
                        GraphicsUnit.Pixel);
                }

                double shortSide = Math.Min(cropW, cropH);
                double pageShort = Math.Min(pageW, pageH);
                double smallThreshold = Math.Max(32.0, Math.Round(0.06 * pageShort));
                double scale = 1.0;
                Bitmap output = cropped;
                if (shortSide < smallThreshold)
                {
                    double pageLong = Math.Max(pageW, pageH);
                    double minLong = Math.Min(256.0, Math.Round(0.25 * pageLong));
                    double cropLong = Math.Max(cropW, cropH);
                    if (cropLong < minLong && cropLong > 0)
                    {
                        scale = minLong / cropLong;
                        int outW = Math.Max(1, (int)Math.Round(cropW * scale));
                        int outH = Math.Max(1, (int)Math.Round(cropH * scale));
                        output = ResizeHighQuality(cropped, outW, outH);
                        if (!ReferenceEquals(output, cropped))
                        {
                            cropped.Dispose();
                        }
                    }
                }

                crop = new CropResult
                {
                    Bitmap = output,
                    Scale = scale,
                    LeftPct = 100.0 * x0 / pageW,
                    TopPct = 100.0 * y0 / pageH,
                    WidthPct = 100.0 * cropW / pageW,
                    HeightPct = 100.0 * cropH / pageH
                };
                return true;
            }
            catch (Exception ex)
            {
                error = "裁切形状失败: " + ex.Message;
                return false;
            }
        }

        private static Bitmap ResizeHighQuality(Bitmap source, int width, int height)
        {
            var dest = new Bitmap(width, height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(dest))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(source, new Rectangle(0, 0, width, height));
            }

            return dest;
        }
    }
}
