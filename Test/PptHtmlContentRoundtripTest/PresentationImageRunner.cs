using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>整稿抽图。前缀 pxi-。</summary>
    internal static class PresentationImageRunner
    {
        public static bool Wants(IList<string> nameFilters)
        {
            if (nameFilters == null)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.IndexOf("pxi-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "pxi", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> filters, bool useWpp)
        {
            Console.WriteLine(useWpp ? "整稿抽图 · WPP" : "整稿抽图 · PPT");
            if (Want(filters, "pxi-neg-word"))
            {
                var fails = new List<string>();
                var stub = new KindStub("pxi-word", ChannelKind.Word);
                if (PresentationHostAdapter.TryExtractPresentationImages(stub, out _, out ToolResult err)
                    || err == null
                    || err.Success
                    || (err.Error ?? "").IndexOf("ppt", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    fails.Add(err == null ? "文字渠道应失败" : (err.Error ?? "成功"));
                }

                Finish(run, "pxi-neg-word", fails);
            }

            if (!UserService.Instance.EnsureEffectiveRootInteractive(null))
            {
                run.CaseFail("pxi-", "工作区未就绪");
                return;
            }

            if (useWpp)
            {
                RunWpp(run, filters);
            }
            else
            {
                RunPpt(run, filters);
            }

            Console.WriteLine("整稿抽图 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static void RunPpt(TestRun run, IList<string> filters)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            string png = Path.Combine(Path.GetTempPath(), "pxi-src.png");
            string broken = Path.Combine(Path.GetTempPath(), "pxi-broken.png");
            try
            {
                WritePng(png);
                WritePng(broken);
                app = new PowerPoint.Application { Visible = Office.MsoTriState.msoTrue };
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch (Exception)
                {
                }

                pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                if (pres.Slides.Count < 1)
                {
                    pres.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                }

                float sw = pres.PageSetup.SlideWidth;
                float sh = pres.PageSetup.SlideHeight;
                PowerPoint.Slide s1 = Blank(pres.Slides[1]);
                Text(s1, "页首句子", 20);
                s1.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 40, 80, sw * 0.2f, sh * 0.2f);
                Text(s1, "页尾句子", 220);

                pres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide s2 = Blank(pres.Slides[2]);
                s2.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 20, 20, sw * 0.08f, sh * 0.08f);
                s2.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 200, 20, sw * 0.10f, sh * 0.10f);

                pres.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide s3 = Blank(pres.Slides[3]);
                PowerPoint.Shape a = Text(s3, "组内甲", 20);
                PowerPoint.Shape pic = s3.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 40, 80, sw * 0.2f, sh * 0.2f);
                PowerPoint.Shape b = Text(s3, "组内乙", 240);
                s3.Shapes.Range(new object[] { a.Name, pic.Name, b.Name }).Group();

                pres.Slides.Add(4, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide s4 = Blank(pres.Slides[4]);
                PowerPoint.Shape chart = s4.Shapes.AddChart(Office.XlChartType.xlColumnClustered, 40, 40, 360, 200);
                chart.Chart.HasTitle = true;
                chart.Chart.ChartTitle.Text = "销量标题";

                pres.Slides.Add(5, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide s5 = Blank(pres.Slides[5]);
                Text(s5, "隐藏页有图", 20);
                s5.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 40, 80, sw * 0.2f, sh * 0.2f);
                s5.SlideShowTransition.Hidden = Office.MsoTriState.msoTrue;

                pres.Slides.Add(6, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Slide s6 = Blank(pres.Slides[6]);
                Text(s6, "坏图页", 20);
                s6.Shapes.AddPicture(png, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 40, 80, sw * 0.2f, sh * 0.2f);

                PptChannel channel = ChannelRegistry.CreateOrGetPpt(pres, "", claimDefaultIfEmpty: true);
                PresentationImageExtract.TestFailExportOrdinal = 5;
                AssertDeck(run, filters, channel.ChannelId, invokeWithoutChannel: true);
            }
            catch (Exception ex)
            {
                run.CaseFail("pxi-", "COM 失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        pres.Saved = Office.MsoTriState.msoTrue;
                        pres.Close();
                    }
                }
                catch (Exception)
                {
                }

                try
                {
                    if (app != null && app.Presentations.Count == 0)
                    {
                        app.Quit();
                    }
                }
                catch (Exception)
                {
                }

                PresentationImageExtract.TestFailExportOrdinal = 0;
                TryDelete(png);
                TryDelete(broken);
            }
        }

        private static void RunWpp(TestRun run, IList<string> filters)
        {
            object pres = null;
            string png = Path.Combine(Path.GetTempPath(), "pxi-src.png");
            string broken = Path.Combine(Path.GetTempPath(), "pxi-broken.png");
            try
            {
                WritePng(png);
                WritePng(broken);
                if (!WppApplicationResolver.TryResolve(out object app, out string error, createIfMissing: true))
                {
                    run.CaseFail("pxi-", error ?? "没有 WPS 演示");
                    return;
                }

                pres = WppCom.Invoke(WppCom.GetProperty(app, "Presentations"), "Add");
                object slides = WppCom.GetProperty(pres, "Slides");
                if (Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0) < 1)
                {
                    WppCom.Invoke(slides, "Add", 1, 12);
                }

                object setup = WppCom.GetProperty(pres, "PageSetup");
                float sw = Convert.ToSingle(WppCom.GetProperty(setup, "SlideWidth"));
                float sh = Convert.ToSingle(WppCom.GetProperty(setup, "SlideHeight"));
                object s1 = WppSlide(slides, 1);
                WppText(s1, "页首句子", 20);
                WppPic(s1, png, false, true, 40, 80, sw * 0.2f, sh * 0.2f);
                WppText(s1, "页尾句子", 220);

                WppCom.Invoke(slides, "Add", 2, 12);
                object s2 = WppCom.GetIndexed(slides, 2);
                WppPic(s2, png, false, true, 20, 20, sw * 0.08f, sh * 0.08f);
                WppPic(s2, png, false, true, 200, 20, sw * 0.10f, sh * 0.10f);

                WppCom.Invoke(slides, "Add", 3, 12);
                object s3 = WppCom.GetIndexed(slides, 3);
                object a = WppText(s3, "组内甲", 20);
                object pic = WppPic(s3, png, false, true, 40, 80, sw * 0.2f, sh * 0.2f);
                object b = WppText(s3, "组内乙", 240);
                object range = WppCom.Invoke(
                    WppCom.GetProperty(s3, "Shapes"),
                    "Range",
                    new object[]
                    {
                        new object[]
                        {
                            Convert.ToString(WppCom.GetProperty(a, "Name")),
                            Convert.ToString(WppCom.GetProperty(pic, "Name")),
                            Convert.ToString(WppCom.GetProperty(b, "Name"))
                        }
                    });
                WppCom.Invoke(range, "Group");

                WppCom.Invoke(slides, "Add", 4, 12);
                object s4 = WppCom.GetIndexed(slides, 4);
                object chart = WppCom.Invoke(
                    WppCom.GetProperty(s4, "Shapes"),
                    "AddChart",
                    (int)Office.XlChartType.xlColumnClustered,
                    40f, 40f, 360f, 200f);
                object chartObj = WppCom.GetProperty(chart, "Chart");
                WppCom.TrySetProperty(chartObj, "HasTitle", true);
                WppCom.TrySetProperty(WppCom.GetProperty(chartObj, "ChartTitle"), "Text", "销量标题");

                WppCom.Invoke(slides, "Add", 5, 12);
                object s5 = WppCom.GetIndexed(slides, 5);
                WppText(s5, "隐藏页有图", 20);
                WppPic(s5, png, false, true, 40, 80, sw * 0.2f, sh * 0.2f);
                WppCom.TrySetProperty(WppCom.GetProperty(s5, "SlideShowTransition"), "Hidden", -1);

                WppCom.Invoke(slides, "Add", 6, 12);
                object s6 = WppCom.GetIndexed(slides, 6);
                WppText(s6, "坏图页", 20);
                WppPic(s6, png, false, true, 40, 80, sw * 0.2f, sh * 0.2f);

                WppChannel channel = ChannelRegistry.CreateOrGetWpp(pres, "", claimDefaultIfEmpty: true);
                PresentationImageExtract.TestFailExportOrdinal = 5;
                AssertDeck(run, filters, channel.ChannelId, invokeWithoutChannel: true);
            }
            catch (Exception ex)
            {
                run.CaseFail("pxi-", "COM 失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        WppCom.TrySetProperty(pres, "Saved", true);
                        WppCom.Invoke(pres, "Close");
                    }
                }
                catch (Exception)
                {
                }

                PresentationImageExtract.TestFailExportOrdinal = 0;
                TryDelete(png);
                TryDelete(broken);
            }
        }

        private static void AssertDeck(TestRun run, IList<string> filters, string channelId, bool invokeWithoutChannel)
        {
            ToolResult first = Invoke(channelId);
            PresentationImageExtract.TestFailExportOrdinal = 0;
            string display = Get(first, "display_contents");
            string page1 = Page(display, 1);
            string page2 = Page(display, 2);
            string page3 = Page(display, 3);
            string page4 = Page(display, 4);
            string page5 = Page(display, 5);
            string page6 = Page(display, 6);
            Match tag = Regex.Match(page1 ?? "", "<image slide_id=\"([0-9]+)\" shape_id=\"(sid[0-9]+-s[0-9]+)\" src=\"([^\"]+)\"/>");

            if (Want(filters, "pxi-place"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                int head = page1.IndexOf("页首句子", StringComparison.Ordinal);
                int image = page1.IndexOf("<image", StringComparison.Ordinal);
                int tail = page1.IndexOf("页尾句子", StringComparison.Ordinal);
                if (!tag.Success || !(head >= 0 && head < image && image < tail))
                {
                    fails.Add("图没有夹在两句之间，或没有路径");
                }
                else if (page1.IndexOf("Left", StringComparison.Ordinal) >= 0)
                {
                    fails.Add("写了坐标");
                }
                else
                {
                    string src = tag.Groups[3].Value.Replace('/', '\\');
                    if (!File.Exists(src) || src.IndexOf("\\curr_docs\\", StringComparison.OrdinalIgnoreCase) < 0
                        || src.IndexOf("_ppt\\image_1.png", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        fails.Add("文件不在 curr_docs 的 image_1.png: " + tag.Groups[3].Value);
                    }
                }

                Finish(run, "pxi-place", fails);
            }

            if (Want(filters, "pxi-small"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                int n = Regex.Matches(page2 ?? "", "<image ").Count;
                if (n != 1)
                {
                    fails.Add("第 2 页应只留下 10% 那张，实际 " + n);
                }

                Finish(run, "pxi-small", fails);
            }

            if (Want(filters, "pxi-json-slim"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                string slides = Slides(first);
                if (slides.IndexOf("src", StringComparison.OrdinalIgnoreCase) >= 0
                    || slides.IndexOf("页首句子", StringComparison.Ordinal) >= 0)
                {
                    fails.Add("slides 抄了正文或路径");
                }

                Finish(run, "pxi-json-slim", fails);
            }

            if (Want(filters, "pxi-overview-mark"))
            {
                var fails = new List<string>();
                ToolResult overview = InvokeOverview(channelId);
                Ok(overview, fails);
                string text = Get(overview, "display_contents");
                if (text.IndexOf(PresentationSlidePlainText.PictureMark, StringComparison.Ordinal) < 0)
                {
                    fails.Add("概览没有「这里有一张图」");
                }

                if (tag.Success && text.IndexOf(tag.Groups[3].Value, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    fails.Add("概览写出了路径");
                }

                Finish(run, "pxi-overview-mark", fails);
            }

            if (Want(filters, "pxi-group"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                if (page3.IndexOf("组内甲", StringComparison.Ordinal) < 0
                    || page3.IndexOf("组内乙", StringComparison.Ordinal) < 0
                    || page3.IndexOf("src=", StringComparison.Ordinal) < 0)
                {
                    fails.Add("组内文字或图没出来");
                }

                Finish(run, "pxi-group", fails);
            }

            if (Want(filters, "pxi-chart"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                if (page4.IndexOf("销量标题", StringComparison.Ordinal) < 0
                    || page4.IndexOf("<image", StringComparison.Ordinal) >= 0)
                {
                    fails.Add("图表不该导出成图");
                }

                Finish(run, "pxi-chart", fails);
            }

            if (Want(filters, "pxi-hidden"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                if (page5.IndexOf("隐藏页有图", StringComparison.Ordinal) < 0
                    || page5.IndexOf("src=", StringComparison.Ordinal) < 0)
                {
                    fails.Add("隐藏页的图没抽到");
                }

                Finish(run, "pxi-hidden", fails);
            }

            if (Want(filters, "pxi-fail"))
            {
                var fails = new List<string>();
                Ok(first, fails);
                if (Get(first, "message") != "抽取成功")
                {
                    fails.Add("整次应仍是抽取成功");
                }

                if (page6.IndexOf(PresentationImageExtract.FailWarning, StringComparison.Ordinal) < 0
                    || page6.IndexOf("src=", StringComparison.Ordinal) >= 0
                    || page6.IndexOf("slide_id=", StringComparison.Ordinal) < 0)
                {
                    fails.Add("坏图页应占位并警告，且没有 src。实际 " + page6);
                }

                if (page1.IndexOf("src=", StringComparison.Ordinal) < 0)
                {
                    fails.Add("其他图的路径被冲掉了");
                }

                Finish(run, "pxi-fail", fails);
            }

            if (Want(filters, "pxi-default") && invokeWithoutChannel)
            {
                var fails = new List<string>();
                ToolResult bare = Invoke(null);
                Ok(bare, fails);
                if (Get(bare, "display_contents").IndexOf("页首句子", StringComparison.Ordinal) < 0)
                {
                    fails.Add("不传 channel_id 没有抽到默认稿");
                }

                Finish(run, "pxi-default", fails);
            }

            if (Want(filters, "pxi-reclear") && tag.Success)
            {
                var fails = new List<string>();
                string dir = Path.GetDirectoryName(tag.Groups[3].Value.Replace('/', '\\'));
                string decoy = Path.Combine(dir, "leftover.png");
                File.WriteAllText(decoy, "x");
                ToolResult second = Invoke(channelId);
                Ok(second, fails);
                if (File.Exists(decoy))
                {
                    fails.Add("再次抽图没有删掉旧文件");
                }

                string[] pngs = Directory.GetFiles(dir, "*.png");
                for (int i = 0; i < pngs.Length; i++)
                {
                    if (!Path.GetFileName(pngs[i]).StartsWith("image_", StringComparison.OrdinalIgnoreCase))
                    {
                        fails.Add("目录里有非 image_N 的文件");
                    }
                }

                if (!File.Exists(Path.Combine(dir, "image_1.png")))
                {
                    fails.Add("没有从 image_1 重写");
                }

                Finish(run, "pxi-reclear", fails);
                try
                {
                    if (Directory.Exists(dir))
                    {
                        Directory.Delete(dir, true);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static ToolResult Invoke(string channelId)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_ExtractPresentationImagesTool.Register(tools, null);
            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            return tools["F_extract_presentation_images"](args).GetAwaiter().GetResult();
        }

        private static ToolResult InvokeOverview(string channelId)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_GetPresentationContentTool.Register(tools, null);
            return tools["F_get_presentation_content"](new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["channel_id"] = channelId
            }).GetAwaiter().GetResult();
        }

        private static PowerPoint.Slide Blank(PowerPoint.Slide slide)
        {
            slide.Layout = PowerPoint.PpSlideLayout.ppLayoutBlank;
            for (int i = slide.Shapes.Count; i >= 1; i--)
            {
                try
                {
                    slide.Shapes[i].Delete();
                }
                catch (Exception)
                {
                }
            }

            return slide;
        }

        private static PowerPoint.Shape Text(PowerPoint.Slide slide, string text, float top)
        {
            PowerPoint.Shape box = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 40, top, 400, 36);
            box.TextFrame.TextRange.Text = text;
            return box;
        }

        private static object WppSlide(object slides, int index)
        {
            object slide = WppCom.GetIndexed(slides, index);
            WppCom.TrySetProperty(slide, "Layout", 12);
            object shapes = WppCom.GetProperty(slide, "Shapes");
            int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count"));
            for (int i = count; i >= 1; i--)
            {
                try
                {
                    WppCom.Invoke(WppCom.Invoke(shapes, "Item", i), "Delete");
                }
                catch (Exception)
                {
                }
            }

            return slide;
        }

        private static object WppText(object slide, string text, float top)
        {
            object box = WppCom.Invoke(WppCom.GetProperty(slide, "Shapes"), "AddTextbox", 1, 40f, top, 400f, 36f);
            WppCom.TrySetProperty(WppCom.GetProperty(WppCom.GetProperty(box, "TextFrame"), "TextRange"), "Text", text);
            return box;
        }

        private static object WppPic(object slide, string path, bool link, bool save, float left, float top, float width, float height)
        {
            return WppCom.Invoke(
                WppCom.GetProperty(slide, "Shapes"),
                "AddPicture",
                path,
                link,
                save,
                left,
                top,
                width,
                height);
        }

        private static void WritePng(string path)
        {
            using (var bmp = new Bitmap(32, 32))
            {
                bmp.SetPixel(0, 0, Color.Blue);
                bmp.Save(path, ImageFormat.Png);
            }
        }

        private static string Page(string display, int index)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            var matches = Regex.Matches(display, @"index=(\d+) slide_id=[^\r\n]*\r?\n");
            for (int i = 0; i < matches.Count; i++)
            {
                if (matches[i].Groups[1].Value != index.ToString())
                {
                    continue;
                }

                int start = matches[i].Index + matches[i].Length;
                int end = i + 1 < matches.Count ? matches[i + 1].Index : display.Length;
                string body = display.Substring(start, end - start);
                int note = body.IndexOf("说明：", StringComparison.Ordinal);
                if (note >= 0)
                {
                    body = body.Substring(0, note);
                }

                return body.Trim();
            }

            return "";
        }

        private static string Get(ToolResult result, string key)
        {
            var data = result != null ? result.Data as Dictionary<string, object> : null;
            if (data == null || !data.TryGetValue(key, out object value) || value == null)
            {
                return "";
            }

            return Convert.ToString(value) ?? "";
        }

        private static string Slides(ToolResult result)
        {
            var data = result != null ? result.Data as Dictionary<string, object> : null;
            if (data == null || !data.TryGetValue("slides", out object slides) || slides == null)
            {
                return "";
            }

            return JsonConvert.SerializeObject(slides);
        }

        private static void Ok(ToolResult result, List<string> fails)
        {
            if (result == null || !result.Success)
            {
                fails.Add(result == null ? "工具不可调" : (result.Error ?? "失败"));
            }
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

        private static bool Want(IList<string> filters, string id)
        {
            if (filters == null || filters.Count == 0)
            {
                return true;
            }

            for (int i = 0; i < filters.Count; i++)
            {
                string f = filters[i] ?? "";
                if (string.Equals(f, "pxi", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "pxi-", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (f.Length > 0
                    && (id.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0
                        || f.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    return true;
                }
            }

            return false;
        }

        private static void Finish(TestRun run, string id, List<string> fails)
        {
            if (fails.Count == 0)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, string.Join("；", fails));
            }
        }

        private sealed class KindStub : IOperationChannel
        {
            public KindStub(string id, ChannelKind kind)
            {
                ChannelId = id;
                Kind = kind;
            }

            public string ChannelId { get; }

            public ChannelKind Kind { get; }
        }
    }
}
