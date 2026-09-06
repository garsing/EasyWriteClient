using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;

namespace WordAddIn1
{
    /// <summary>
    /// 截图压缩：与现网 PageCaptureHelper 同一套上限与阶梯。
    /// </summary>
    internal static class ImageCaptureCompressor
    {
        /// <summary>视觉轮按 base64 计 token，必须先缩小。200KB / 长边 800 足够看版式。</summary>
        public const int MaxImageBytes = 204800;
        public const int MaxImageDimension = 800;
        public const int ContactSheetMaxPages = 9;
        public const int ContactSheetCellLongEdge = 320;
        private const int JpegQualityStart = 85;

        internal sealed class ContactSheetTile
        {
            public int Index { get; set; }
            public string SlideId { get; set; }
            public byte[] ImageBytes { get; set; }
        }

        public static void FitExportPixelSize(float slideWidth, float slideHeight, out int width, out int height)
        {
            FitExportPixelSize(slideWidth, slideHeight, MaxImageDimension, out width, out height);
        }

        public static void FitExportPixelSize(
            float slideWidth,
            float slideHeight,
            int maxLongEdge,
            out int width,
            out int height)
        {
            int cap = maxLongEdge > 0 ? maxLongEdge : MaxImageDimension;
            if (slideWidth <= 0f || slideHeight <= 0f)
            {
                width = cap;
                height = Math.Max(1, cap * 9 / 16);
                return;
            }

            if (slideWidth >= slideHeight)
            {
                width = cap;
                height = Math.Max(1, (int)Math.Round(cap * (double)slideHeight / slideWidth));
                return;
            }

            height = cap;
            width = Math.Max(1, (int)Math.Round(cap * (double)slideWidth / slideHeight));
        }

        internal sealed class CompressedImage
        {
            public byte[] Bytes { get; set; }
            public string Format { get; set; }
            public int Width { get; set; }
            public int Height { get; set; }
        }

        public static CompressedImage Compress(Bitmap source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            using (var pngCopy = new Bitmap(source))
            {
                byte[] pngBytes = EncodePng(pngCopy);
                if (pngBytes.Length <= MaxImageBytes &&
                    Math.Max(pngCopy.Width, pngCopy.Height) <= MaxImageDimension)
                {
                    return new CompressedImage
                    {
                        Bytes = pngBytes,
                        Format = "png",
                        Width = pngCopy.Width,
                        Height = pngCopy.Height
                    };
                }
            }

            int quality = JpegQualityStart;
            double scale = 1.0;
            Bitmap working = new Bitmap(source);

            try
            {
                while (true)
                {
                    int width = Math.Max(1, (int)Math.Round(working.Width * scale));
                    int height = Math.Max(1, (int)Math.Round(working.Height * scale));
                    if (Math.Max(width, height) > MaxImageDimension)
                    {
                        double fit = (double)MaxImageDimension / Math.Max(working.Width, working.Height);
                        width = Math.Max(1, (int)Math.Round(working.Width * fit));
                        height = Math.Max(1, (int)Math.Round(working.Height * fit));
                    }

                    using (var resized = ResizeBitmap(working, width, height))
                    {
                        byte[] jpegBytes = EncodeJpeg(resized, quality);
                        if (jpegBytes.Length <= MaxImageBytes)
                        {
                            return new CompressedImage
                            {
                                Bytes = jpegBytes,
                                Format = "jpeg",
                                Width = resized.Width,
                                Height = resized.Height
                            };
                        }
                    }

                    if (quality > 55)
                    {
                        quality -= 10;
                        continue;
                    }

                    if (scale > 0.35)
                    {
                        scale *= 0.85;
                        quality = JpegQualityStart;
                        continue;
                    }

                    throw new InvalidOperationException(
                        "截图压缩后仍超过大小限制(" + MaxImageBytes + " bytes)");
                }
            }
            finally
            {
                working.Dispose();
            }
        }

        public static CompressedImage ComposeContactSheet(IReadOnlyList<ContactSheetTile> tiles)
        {
            if (tiles == null || tiles.Count == 0)
            {
                throw new ArgumentException("宫格没有页");
            }

            if (tiles.Count > ContactSheetMaxPages)
            {
                throw new ArgumentException("一次最多 " + ContactSheetMaxPages + " 页");
            }

            const int colsMax = 3;
            const int gap = 8;
            const int labelH = 22;
            const int margin = 8;
            const int cellW = ContactSheetCellLongEdge;

            int n = tiles.Count;
            int cols = Math.Min(colsMax, n);
            int rows = (n + cols - 1) / cols;
            int cellImgH = cellW * 9 / 16;
            var decoded = new List<Bitmap>(n);
            try
            {
                for (int i = 0; i < n; i++)
                {
                    Bitmap bmp = DecodeBitmap(tiles[i] == null ? null : tiles[i].ImageBytes);
                    decoded.Add(bmp);
                    if (i == 0 && bmp != null && bmp.Width > 0)
                    {
                        cellImgH = Math.Max(1, (int)Math.Round(cellW * (double)bmp.Height / bmp.Width));
                    }
                }

                int cellH = labelH + cellImgH;
                int width = margin * 2 + cols * cellW + (cols - 1) * gap;
                int height = margin * 2 + rows * cellH + (rows - 1) * gap;
                using (var canvas = new Bitmap(width, height))
                using (var graphics = Graphics.FromImage(canvas))
                using (var font = CreateLabelFont())
                using (var imgBg = new SolidBrush(Color.FromArgb(32, 36, 42)))
                using (var bar = new SolidBrush(Color.FromArgb(27, 58, 75)))
                using (var white = new SolidBrush(Color.White))
                {
                    graphics.Clear(Color.FromArgb(236, 236, 236));
                    graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    for (int i = 0; i < n; i++)
                    {
                        int row = i / cols;
                        int col = i % cols;
                        int x = margin + col * (cellW + gap);
                        int y = margin + row * (cellH + gap);
                        graphics.FillRectangle(bar, x, y, cellW, labelH);
                        string slideId = tiles[i] == null ? "" : (tiles[i].SlideId ?? "");
                        int index = tiles[i] == null ? i + 1 : tiles[i].Index;
                        string label = string.IsNullOrEmpty(slideId)
                            ? index.ToString()
                            : index + " · " + slideId;
                        graphics.DrawString(
                            label,
                            font,
                            white,
                            new RectangleF(x + 4, y + 2, cellW - 8, labelH - 2));
                        int imgY = y + labelH;
                        graphics.FillRectangle(imgBg, x, imgY, cellW, cellImgH);
                        Bitmap src = decoded[i];
                        if (src != null && src.Width > 0 && src.Height > 0)
                        {
                            double scale = Math.Min(
                                (double)cellW / src.Width,
                                (double)cellImgH / src.Height);
                            int dw = Math.Max(1, (int)Math.Round(src.Width * scale));
                            int dh = Math.Max(1, (int)Math.Round(src.Height * scale));
                            int dx = x + (cellW - dw) / 2;
                            int dy = imgY + (cellImgH - dh) / 2;
                            graphics.DrawImage(src, dx, dy, dw, dh);
                        }
                    }

                    return Compress(canvas);
                }
            }
            finally
            {
                for (int i = 0; i < decoded.Count; i++)
                {
                    if (decoded[i] != null)
                    {
                        decoded[i].Dispose();
                    }
                }
            }
        }

        public static CompressedImage CompressThumb(Bitmap source, int maxLongEdge = 320, int quality = 80)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            int longEdge = Math.Max(source.Width, source.Height);
            int width = source.Width;
            int height = source.Height;
            if (longEdge > maxLongEdge)
            {
                double scale = (double)maxLongEdge / longEdge;
                width = Math.Max(1, (int)Math.Round(source.Width * scale));
                height = Math.Max(1, (int)Math.Round(source.Height * scale));
            }

            using (var resized = (width == source.Width && height == source.Height)
                ? new Bitmap(source)
                : ResizeBitmap(source, width, height))
            {
                int q = quality;
                byte[] payload = EncodeJpeg(resized, q);
                while (payload.Length > MaxImageBytes && q > 40)
                {
                    q -= 10;
                    payload = EncodeJpeg(resized, q);
                }

                return new CompressedImage
                {
                    Bytes = payload,
                    Format = "jpeg",
                    Width = resized.Width,
                    Height = resized.Height
                };
            }
        }

        public static CompressedImage CompressFile(string imagePath)
        {
            using (var bitmap = new Bitmap(imagePath))
            {
                return Compress(bitmap);
            }
        }

        public static CompressedImage CompressBytes(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                throw new InvalidOperationException("图片字节为空");
            }

            using (var stream = new MemoryStream(bytes, writable: false))
            using (var bitmap = new Bitmap(stream))
            {
                return Compress(bitmap);
            }
        }

        /// <summary>
        /// CopyPicture 贴进空 Chart 时经常只导出白底淡框。用深色像素占比判断。
        /// </summary>
        public static bool LooksMostlyBlankFile(string imagePath)
        {
            if (string.IsNullOrEmpty(imagePath) || !File.Exists(imagePath))
            {
                return true;
            }

            using (var bitmap = new Bitmap(imagePath))
            {
                int dark = 0;
                int total = 0;
                int stepX = Math.Max(1, bitmap.Width / 80);
                int stepY = Math.Max(1, bitmap.Height / 80);
                for (int y = 0; y < bitmap.Height; y += stepY)
                {
                    for (int x = 0; x < bitmap.Width; x += stepX)
                    {
                        Color c = bitmap.GetPixel(x, y);
                        total++;
                        if ((0.299 * c.R) + (0.587 * c.G) + (0.114 * c.B) < 160)
                        {
                            dark++;
                        }
                    }
                }

                return total == 0 || (dark * 200) < total;
            }
        }

        private static Bitmap DecodeBitmap(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                return null;
            }

            using (var stream = new MemoryStream(bytes, writable: false))
            using (var loaded = new Bitmap(stream))
            {
                return new Bitmap(loaded);
            }
        }

        private static Font CreateLabelFont()
        {
            try
            {
                return new Font("Microsoft YaHei", 11f, FontStyle.Bold, GraphicsUnit.Pixel);
            }
            catch (Exception)
            {
                return new Font(FontFamily.GenericSansSerif, 11f, FontStyle.Bold, GraphicsUnit.Pixel);
            }
        }

        private static Bitmap ResizeBitmap(Bitmap source, int width, int height)
        {
            var target = new Bitmap(width, height);
            using (var graphics = Graphics.FromImage(target))
            {
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.DrawImage(source, 0, 0, width, height);
            }

            return target;
        }

        private static byte[] EncodePng(Bitmap bitmap)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }

        private static byte[] EncodeJpeg(Bitmap bitmap, int quality)
        {
            var codec = GetJpegCodec();
            using (var stream = new MemoryStream())
            {
                if (codec == null)
                {
                    bitmap.Save(stream, ImageFormat.Jpeg);
                    return stream.ToArray();
                }

                var encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = new EncoderParameter(Encoder.Quality, (long)quality);
                bitmap.Save(stream, codec, encoderParams);
                return stream.ToArray();
            }
        }

        private static ImageCodecInfo GetJpegCodec()
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            foreach (var codec in codecs)
            {
                if (codec.FormatID == ImageFormat.Jpeg.Guid)
                {
                    return codec;
                }
            }

            return null;
        }
    }
}
