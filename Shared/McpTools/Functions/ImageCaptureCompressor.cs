using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace WordAddIn1
{
    /// <summary>
    /// 截图压缩：与现网 PageCaptureHelper 同一套上限与阶梯。
    /// </summary>
    internal static class ImageCaptureCompressor
    {
        public const int MaxImageBytes = 4194304;
        public const int MaxImageDimension = 2048;
        private const int JpegQualityStart = 85;

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

        public static CompressedImage CompressFile(string imagePath)
        {
            using (var bitmap = new Bitmap(imagePath))
            {
                return Compress(bitmap);
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
