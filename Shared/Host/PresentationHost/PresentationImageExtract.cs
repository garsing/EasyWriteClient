using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using WordAddIn1.OpenFiles;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 一次导出已打开演示文稿里够大的图片。文字顺序与概览相同，图片标签插在原处。
    /// </summary>
    internal static class PresentationImageExtract
    {
        public const string FailWarning = "警告：抽取失败";

        /// <summary>测试用。第 N 次真正导出（1 起）改为失败占位。0 表示关闭。</summary>
        internal static int TestFailExportOrdinal;

        private const int PpShapeFormatPng = 2;

        private const double MinArea = 1.0;

        public static bool TryPowerPoint(
            PowerPoint.Presentation presentation,
            string channelId,
            out PresentationContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            string name;
            string path;
            int count;
            double slideWidth;
            double slideHeight;
            try
            {
                name = presentation.Name ?? "";
                path = WordAddIn1.PptChannel.TryReadFullName(presentation) ?? "";
                count = presentation.Slides.Count;
                PowerPoint.PageSetup setup = presentation.PageSetup;
                slideWidth = setup.SlideWidth;
                slideHeight = setup.SlideHeight;
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            if (!TryResetDirectory(name, out string directory, out error))
            {
                return false;
            }

            var sink = new Sink(directory, slideWidth, slideHeight);
            List<PresentationSlideInfo> slides = ReadPowerPoint(
                presentation, count, sink, out bool truncated, out string truncatedReason, out error);
            if (slides == null)
            {
                return false;
            }

            result = Finish(channelId, "ppt", name, path, count, slides, truncated, truncatedReason);
            return true;
        }

        public static bool TryWpp(
            object presentation,
            string channelId,
            out PresentationContentResult result,
            out string error)
        {
            result = null;
            error = null;
            if (presentation == null)
            {
                error = "渠道对应的演示文稿已关闭";
                return false;
            }

            string name;
            string path;
            int count;
            double slideWidth;
            double slideHeight;
            object slidesCollection;
            try
            {
                name = WppCom.TryReadName(presentation) ?? "";
                path = WppCom.TryReadFullName(presentation) ?? "";
                slidesCollection = WppCom.GetProperty(presentation, "Slides");
                count = Convert.ToInt32(WppCom.GetProperty(slidesCollection, "Count"));
                object setup = WppCom.GetProperty(presentation, "PageSetup");
                slideWidth = Convert.ToDouble(WppCom.GetProperty(setup, "SlideWidth"));
                slideHeight = Convert.ToDouble(WppCom.GetProperty(setup, "SlideHeight"));
            }
            catch (Exception ex)
            {
                error = "COM 不可用: " + ex.Message;
                return false;
            }

            if (!TryResetDirectory(name, out string directory, out error))
            {
                return false;
            }

            var sink = new Sink(directory, slideWidth, slideHeight);
            List<PresentationSlideInfo> slides = ReadWpp(
                slidesCollection, count, sink, out bool truncated, out string truncatedReason, out error);
            if (slides == null)
            {
                return false;
            }

            result = Finish(channelId, "wpp", name, path, count, slides, truncated, truncatedReason);
            return true;
        }

        public static string BuildDisplay(PresentationContentResult result)
        {
            if (result == null)
            {
                return "";
            }

            var sb = new StringBuilder();
            sb.AppendLine("演示文稿：" + (result.Name ?? ""));
            sb.AppendLine("路径：" + (string.IsNullOrEmpty(result.Path) ? "（未保存）" : result.Path));
            sb.AppendLine(
                "kind=" + (result.Kind ?? "")
                + " channel_id=" + (result.ChannelId ?? ""));
            sb.AppendLine("页数：" + result.SlideCount);
            if (result.Truncated)
            {
                sb.AppendLine(
                    "truncated=true"
                    + (string.IsNullOrEmpty(result.TruncatedReason)
                        ? ""
                        : " reason=" + result.TruncatedReason));
            }

            sb.AppendLine();
            if (result.Slides == null || result.Slides.Count == 0)
            {
                sb.AppendLine("（无幻灯片）");
                return sb.ToString().TrimEnd();
            }

            foreach (PresentationSlideInfo slide in result.Slides)
            {
                if (slide == null)
                {
                    continue;
                }

                sb.Append("index=").Append(slide.Index);
                sb.Append(" slide_id=").Append(slide.SlideId ?? "");
                sb.AppendLine();
                string body = string.IsNullOrWhiteSpace(slide.PageText)
                    ? "（无文字）"
                    : slide.PageText.Trim();
                sb.AppendLine(body);
                sb.AppendLine();
            }

            sb.Append("说明：路径在 src。失败的位置有「").Append(FailWarning).Append("」，没有 src。");
            return sb.ToString().TrimEnd();
        }

        private static bool TryResetDirectory(string presentationName, out string directory, out string error)
        {
            directory = null;
            error = null;
            string root = UserService.Instance == null ? null : UserService.Instance.WorkspaceRootEffective;
            if (string.IsNullOrWhiteSpace(root))
            {
                error = "工作区根不可用，无法导出图片";
                return false;
            }

            string stem = CurrDocsImagePaths.SanitizeStem(presentationName);
            directory = Path.Combine(root, "curr_docs", stem + "_ppt");
            try
            {
                if (Directory.Exists(directory))
                {
                    Directory.Delete(directory, true);
                }

                Directory.CreateDirectory(directory);
            }
            catch (Exception ex)
            {
                directory = null;
                error = "无法准备图片目录: " + ex.Message;
                return false;
            }

            return true;
        }

        private static List<PresentationSlideInfo> ReadPowerPoint(
            PowerPoint.Presentation presentation,
            int count,
            Sink sink,
            out bool truncated,
            out string truncatedReason,
            out string error)
        {
            truncated = false;
            truncatedReason = null;
            error = null;
            int limit = Math.Min(count, PresentationContentResult.MaxSlides);
            if (count > PresentationContentResult.MaxSlides)
            {
                truncated = true;
                truncatedReason = "页数超过 " + PresentationContentResult.MaxSlides
                    + "，仅导出前 " + PresentationContentResult.MaxSlides + " 页";
            }

            var slides = new List<PresentationSlideInfo>();
            try
            {
                for (int i = 1; i <= limit; i++)
                {
                    PowerPoint.Slide slide = presentation.Slides[i];
                    string slideId = slide.SlideID.ToString(CultureInfo.InvariantCulture);
                    slides.Add(new PresentationSlideInfo
                    {
                        Index = i,
                        SlideId = slideId,
                        PageText = PresentationSlidePlainText.CollectPowerPoint(
                            slide,
                            shape => sink.TakePowerPoint(shape, slideId))
                    });
                }
            }
            catch (Exception ex)
            {
                error = "读取幻灯片失败: " + ex.Message;
                return null;
            }

            return slides;
        }

        private static List<PresentationSlideInfo> ReadWpp(
            object slidesCollection,
            int count,
            Sink sink,
            out bool truncated,
            out string truncatedReason,
            out string error)
        {
            truncated = false;
            truncatedReason = null;
            error = null;
            int limit = Math.Min(count, PresentationContentResult.MaxSlides);
            if (count > PresentationContentResult.MaxSlides)
            {
                truncated = true;
                truncatedReason = "页数超过 " + PresentationContentResult.MaxSlides
                    + "，仅导出前 " + PresentationContentResult.MaxSlides + " 页";
            }

            var slides = new List<PresentationSlideInfo>();
            try
            {
                for (int i = 1; i <= limit; i++)
                {
                    object slide = WppCom.GetIndexed(slidesCollection, i);
                    string slideId = Convert.ToString(
                        WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId")) ?? "";
                    slides.Add(new PresentationSlideInfo
                    {
                        Index = i,
                        SlideId = slideId,
                        PageText = PresentationSlidePlainText.CollectWpp(
                            slide,
                            shape => sink.TakeWpp(shape, slideId))
                    });
                }
            }
            catch (Exception ex)
            {
                error = "读取幻灯片失败: " + ex.Message;
                return null;
            }

            return slides;
        }

        private static PresentationContentResult Finish(
            string channelId,
            string kind,
            string name,
            string path,
            int count,
            List<PresentationSlideInfo> slides,
            bool truncated,
            string truncatedReason)
        {
            string saved = path ?? "";
            try
            {
                if (!Path.IsPathRooted(saved))
                {
                    saved = "";
                }
            }
            catch (Exception)
            {
                saved = "";
            }

            return new PresentationContentResult
            {
                ChannelId = channelId ?? "",
                Kind = kind,
                Name = name ?? "",
                Path = saved,
                SlideCount = count,
                Slides = slides,
                Truncated = truncated,
                TruncatedReason = truncatedReason
            };
        }

        private static bool ShouldExport(double width, double height, double slideWidth, double slideHeight)
        {
            if (slideWidth <= 0 || slideHeight <= 0 || width <= 0 || height <= 0)
            {
                return true;
            }

            double area = (width / slideWidth * 100.0) * (height / slideHeight * 100.0) / 100.0;
            if (area <= 0)
            {
                return true;
            }

            return area + 0.0000001 >= MinArea;
        }

        private sealed class Sink
        {
            private readonly string _directory;

            private readonly double _slideWidth;

            private readonly double _slideHeight;

            private int _next = 1;

            private int _attempt;

            public Sink(string directory, double slideWidth, double slideHeight)
            {
                _directory = directory;
                _slideWidth = slideWidth;
                _slideHeight = slideHeight;
            }

            public string TakePowerPoint(PowerPoint.Shape shape, string slideId)
            {
                if (shape == null || !ShouldExport(Safe(() => shape.Width), Safe(() => shape.Height), _slideWidth, _slideHeight))
                {
                    return "";
                }

                string shapeId = "";
                try
                {
                    shapeId = PptShapeId.FormatShape(slideId, shape.Id);
                }
                catch (Exception)
                {
                }

                return Write(slideId, shapeId, path =>
                {
                    shape.Export(path, (PowerPoint.PpShapeFormat)PpShapeFormatPng);
                });
            }

            public string TakeWpp(object shape, string slideId)
            {
                if (shape == null)
                {
                    return "";
                }

                double width = Safe(() => Convert.ToDouble(WppCom.GetProperty(shape, "Width")));
                double height = Safe(() => Convert.ToDouble(WppCom.GetProperty(shape, "Height")));
                if (!ShouldExport(width, height, _slideWidth, _slideHeight))
                {
                    return "";
                }

                string shapeId = "";
                try
                {
                    int comId = Convert.ToInt32(WppCom.GetProperty(shape, "Id"));
                    shapeId = PptShapeId.FormatShape(slideId, comId);
                }
                catch (Exception)
                {
                }

                return Write(slideId, shapeId, path =>
                {
                    WppCom.Invoke(shape, "Export", path, PpShapeFormatPng);
                });
            }

            private string Write(string slideId, string shapeId, Action<string> export)
            {
                _attempt++;
                if (TestFailExportOrdinal > 0 && _attempt == TestFailExportOrdinal)
                {
                    return FailTag(slideId, shapeId);
                }

                string finalPath = Path.Combine(
                    _directory,
                    "image_" + _next.ToString(CultureInfo.InvariantCulture) + ".png");
                try
                {
                    export(finalPath);
                    if (!File.Exists(finalPath))
                    {
                        return FailTag(slideId, shapeId);
                    }

                    _next++;
                    string src = Path.GetFullPath(finalPath).Replace('\\', '/');
                    return "<image slide_id=\"" + Attr(slideId)
                        + "\" shape_id=\"" + Attr(shapeId)
                        + "\" src=\"" + Attr(src) + "\"/>";
                }
                catch (Exception)
                {
                    TryDelete(finalPath);
                    return FailTag(slideId, shapeId);
                }
            }

            private static string FailTag(string slideId, string shapeId)
            {
                return "<image slide_id=\"" + Attr(slideId)
                    + "\" shape_id=\"" + Attr(shapeId)
                    + "\"/>" + "\n" + FailWarning;
            }

            private static double Safe(Func<double> read)
            {
                try
                {
                    return read();
                }
                catch (Exception)
                {
                    return 0;
                }
            }

            private static string Attr(string value)
            {
                return (value ?? "")
                    .Replace("&", "&amp;")
                    .Replace("\"", "&quot;")
                    .Replace("<", "&lt;");
            }

            private static void TryDelete(string path)
            {
                try
                {
                    if (!string.IsNullOrEmpty(path) && File.Exists(path))
                    {
                        File.Delete(path);
                    }
                }
                catch (Exception)
                {
                }
            }
        }
    }
}
