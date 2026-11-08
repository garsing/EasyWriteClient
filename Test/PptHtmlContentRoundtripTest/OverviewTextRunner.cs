using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
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
    /// <summary>概览每页文字。前缀 ovt-。不并进已绿的 ov-。</summary>
    internal static class OverviewTextRunner
    {
        private const string LongSentence =
            "跨应用原生操控、办公场景专项调优、多源数据赋能，逐项破解工具割裂、性能不足、内容不专业";

        private const string ChartTitle = "销量标题";

        private const string ChartStamp = "987654";

        private static string ChartStampError;

        public static bool Wants(IList<string> nameFilters)
        {
            if (nameFilters == null)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.IndexOf("ovt-", StringComparison.OrdinalIgnoreCase) >= 0
                    || string.Equals(f, "ovt", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            Console.WriteLine(useWpp
                ? "概览每页文字 · WPP"
                : "概览每页文字 · PPT");
            if (useWpp)
            {
                RunWpp(run, nameFilters);
            }
            else
            {
                RunPpt(run, nameFilters);
            }

            Console.WriteLine("概览每页文字 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
        }

        private static void RunPpt(TestRun run, IList<string> filters)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            string png = Path.Combine(Path.GetTempPath(), "ovt-mark.png");
            try
            {
                WritePng(png);
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

                pres.Slides[1].Layout = PowerPoint.PpSlideLayout.ppLayoutBlank;
                ClearShapes(pres.Slides[1]);
                PowerPoint.Shape box = pres.Slides[1].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, 40f, 30f, 640f, 80f);
                box.TextFrame.TextRange.Text = LongSentence;

                pres.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Shape table = pres.Slides[2].Shapes.AddTable(2, 2, 40f, 40f, 360f, 120f);
                table.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = "甲格";
                table.Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = "乙格";
                table.Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = "丙格";
                table.Table.Cell(2, 2).Shape.TextFrame.TextRange.Text = "丁格";

                pres.Slides.Add(3, PowerPoint.PpSlideLayout.ppLayoutBlank);
                pres.Slides[3].Shapes.AddPicture(
                    png,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    40f, 40f, 64f, 64f);

                pres.Slides.Add(4, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Shape chart = pres.Slides[4].Shapes.AddChart(
                    Office.XlChartType.xlColumnClustered, 40f, 40f, 360f, 220f);
                chart.Chart.HasTitle = true;
                chart.Chart.ChartTitle.Text = ChartTitle;
                try
                {
                    StampChartData(chart.Chart);
                    ChartStampError = null;
                }
                catch (Exception ex)
                {
                    ChartStampError = ex.Message;
                }

                pres.Slides.Add(5, PowerPoint.PpSlideLayout.ppLayoutBlank);
                PowerPoint.Shape left = pres.Slides[5].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, 40f, 40f, 160f, 40f);
                left.TextFrame.TextRange.Text = "组内甲";
                PowerPoint.Shape right = pres.Slides[5].Shapes.AddTextbox(
                    Office.MsoTextOrientation.msoTextOrientationHorizontal, 220f, 40f, 160f, 40f);
                right.TextFrame.TextRange.Text = "组内乙";
                pres.Slides[5].Shapes.Range(new object[] { left.Name, right.Name }).Group();

                pres.Slides.Add(6, PowerPoint.PpSlideLayout.ppLayoutBlank);
                ClearShapes(pres.Slides[6]);

                PptChannel channel = ChannelRegistry.CreateOrGetPpt(pres, "", claimDefaultIfEmpty: false);
                ToolResult first = Invoke(channel.ChannelId);
                AssertBuilt(run, filters, first, png, channel.ChannelId);

                box.TextFrame.TextRange.Text = "改后句子乙";
                ToolResult second = Invoke(channel.ChannelId);
                if (Want(filters, "ovt-unsaved-edit"))
                {
                    var fails = new List<string>();
                    string page = PageBody(GetDisplay(second), 1);
                    if (page.IndexOf("改后句子乙", StringComparison.Ordinal) < 0)
                    {
                        fails.Add("未保存的新字未出现");
                    }

                    if (page.IndexOf(LongSentence, StringComparison.Ordinal) >= 0)
                    {
                        fails.Add("改前的长句还在");
                    }

                    Finish(run, "ovt-unsaved-edit", fails);
                }
            }
            catch (Exception ex)
            {
                run.CaseFail("ovt-", "COM 失败: " + ex.Message);
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

                try
                {
                    if (File.Exists(png))
                    {
                        File.Delete(png);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static void RunWpp(TestRun run, IList<string> filters)
        {
            object app = null;
            object pres = null;
            string png = Path.Combine(Path.GetTempPath(), "ovt-mark.png");
            string step = "写图";
            try
            {
                WritePng(png);
                step = "打开 WPS";
                if (!WppApplicationResolver.TryResolve(out app, out string error, createIfMissing: true))
                {
                    run.CaseFail("ovt-", error ?? "没有 WPS 演示");
                    return;
                }

                step = "新建文稿";
                object presentations = WppCom.GetProperty(app, "Presentations");
                pres = WppCom.Invoke(presentations, "Add");
                object slides = WppCom.GetProperty(pres, "Slides");
                int slideCount = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                if (slideCount < 1)
                {
                    try
                    {
                        WppCom.Invoke(slides, "Add", 1, 12);
                    }
                    catch (Exception)
                    {
                        WppCom.Invoke(slides, "Add", 1, 1);
                    }
                }

                step = "第 1 页文本框";
                object first = WppCom.GetIndexed(slides, 1);
                TrySetLayoutBlank(first);
                ClearWppShapes(first);
                object box = WppCom.Invoke(WppCom.GetProperty(first, "Shapes"), "AddTextbox", 1, 40f, 30f, 640f, 80f);
                WppCom.TrySetProperty(WppCom.GetProperty(WppCom.GetProperty(box, "TextFrame"), "TextRange"), "Text", LongSentence);

                step = "第 2 页表格";
                WppCom.Invoke(slides, "Add", 2, 12);
                object slide2 = WppCom.GetIndexed(slides, 2);
                object table = WppCom.Invoke(WppCom.GetProperty(slide2, "Shapes"), "AddTable", 2, 2, 40f, 40f, 360f, 120f);
                SetWppCell(table, 1, 1, "甲格");
                SetWppCell(table, 1, 2, "乙格");
                SetWppCell(table, 2, 1, "丙格");
                SetWppCell(table, 2, 2, "丁格");

                step = "第 3 页图片";
                WppCom.Invoke(slides, "Add", 3, 12);
                object slide3 = WppCom.GetIndexed(slides, 3);
                WppCom.Invoke(
                    WppCom.GetProperty(slide3, "Shapes"),
                    "AddPicture",
                    png,
                    false,
                    true,
                    40f,
                    40f,
                    64f,
                    64f);

                step = "第 4 页图表";
                WppCom.Invoke(slides, "Add", 4, 12);
                object slide4 = WppCom.GetIndexed(slides, 4);
                object chart = WppCom.Invoke(
                    WppCom.GetProperty(slide4, "Shapes"),
                    "AddChart",
                    (int)Office.XlChartType.xlColumnClustered,
                    40f,
                    40f,
                    360f,
                    220f);
                object chartObj = WppCom.GetProperty(chart, "Chart");
                WppCom.TrySetProperty(chartObj, "HasTitle", true);
                WppCom.TrySetProperty(WppCom.GetProperty(chartObj, "ChartTitle"), "Text", ChartTitle);
                try
                {
                    StampChartData(chartObj);
                    ChartStampError = null;
                }
                catch (Exception ex)
                {
                    ChartStampError = ex.Message;
                }

                step = "第 5 页组合";
                WppCom.Invoke(slides, "Add", 5, 12);
                object slide5 = WppCom.GetIndexed(slides, 5);
                object shapes5 = WppCom.GetProperty(slide5, "Shapes");
                object left = WppCom.Invoke(shapes5, "AddTextbox", 1, 40f, 40f, 160f, 40f);
                object right = WppCom.Invoke(shapes5, "AddTextbox", 1, 220f, 40f, 160f, 40f);
                WppCom.TrySetProperty(WppCom.GetProperty(WppCom.GetProperty(left, "TextFrame"), "TextRange"), "Text", "组内甲");
                WppCom.TrySetProperty(WppCom.GetProperty(WppCom.GetProperty(right, "TextFrame"), "TextRange"), "Text", "组内乙");
                object range = WppCom.Invoke(
                    shapes5,
                    "Range",
                    new object[]
                    {
                        new object[]
                        {
                            Convert.ToString(WppCom.GetProperty(left, "Name")),
                            Convert.ToString(WppCom.GetProperty(right, "Name"))
                        }
                    });
                WppCom.Invoke(range, "Group");

                step = "第 6 页空白";
                WppCom.Invoke(slides, "Add", 6, 12);
                ClearWppShapes(WppCom.GetIndexed(slides, 6));

                step = "读概览";
                WppChannel channel = ChannelRegistry.CreateOrGetWpp(pres, "", claimDefaultIfEmpty: false);
                ToolResult firstResult = Invoke(channel.ChannelId);
                AssertBuilt(run, filters, firstResult, png, channel.ChannelId);

                WppCom.TrySetProperty(
                    WppCom.GetProperty(WppCom.GetProperty(box, "TextFrame"), "TextRange"),
                    "Text",
                    "改后句子乙");
                ToolResult second = Invoke(channel.ChannelId);
                if (Want(filters, "ovt-unsaved-edit"))
                {
                    var fails = new List<string>();
                    string page = PageBody(GetDisplay(second), 1);
                    if (page.IndexOf("改后句子乙", StringComparison.Ordinal) < 0)
                    {
                        fails.Add("未保存的新字未出现");
                    }

                    if (page.IndexOf(LongSentence, StringComparison.Ordinal) >= 0)
                    {
                        fails.Add("改前的长句还在");
                    }

                    Finish(run, "ovt-unsaved-edit", fails);
                }
            }
            catch (Exception ex)
            {
                string detail = ex.Message;
                if (ex.InnerException != null)
                {
                    detail += " / " + ex.InnerException.Message;
                }

                run.CaseFail("ovt-", "COM 失败（" + step + "）: " + detail);
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

                try
                {
                    if (File.Exists(png))
                    {
                        File.Delete(png);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static void AssertBuilt(
            TestRun run,
            IList<string> filters,
            ToolResult result,
            string pngPath,
            string channelId)
        {
            string display = GetDisplay(result);
            string slidesJson = SlidesJson(result);
            if (Want(filters, "ovt-textbox-full"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 1);
                if (page.IndexOf(LongSentence, StringComparison.Ordinal) < 0)
                {
                    fails.Add("长句不完整");
                }

                if (page.IndexOf("slide_id=", StringComparison.Ordinal) < 0 && display.IndexOf("slide_id=", StringComparison.Ordinal) < 0)
                {
                    fails.Add("缺 slide_id");
                }

                if (display.IndexOf("<section", StringComparison.OrdinalIgnoreCase) >= 0
                    || display.IndexOf("data-src", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    fails.Add("概览含约定 HTML");
                }

                Finish(run, "ovt-textbox-full", fails);
            }

            if (Want(filters, "ovt-table-row"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 2);
                if (page.IndexOf("甲格 乙格", StringComparison.Ordinal) < 0
                    || page.IndexOf("丙格 丁格", StringComparison.Ordinal) < 0)
                {
                    fails.Add("表格行不是「甲格 乙格 / 丙格 丁格」，实际 " + page);
                }

                Finish(run, "ovt-table-row", fails);
            }

            if (Want(filters, "ovt-picture-mark"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 3);
                if (page.IndexOf(PresentationSlidePlainText.PictureMark, StringComparison.Ordinal) < 0)
                {
                    fails.Add("没有「这里有一张图」");
                }

                if (display.IndexOf(pngPath, StringComparison.OrdinalIgnoreCase) >= 0
                    || display.IndexOf("ovt-mark.png", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    fails.Add("概览写出了图片路径");
                }

                Finish(run, "ovt-picture-mark", fails);
            }

            if (Want(filters, "ovt-chart-title"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 4);
                if (page.IndexOf(ChartTitle, StringComparison.Ordinal) < 0)
                {
                    fails.Add("没有图表标题");
                }

                if (page.IndexOf(PresentationSlidePlainText.PictureMark, StringComparison.Ordinal) >= 0)
                {
                    fails.Add("图表被标成图片");
                }

                if (!string.IsNullOrEmpty(ChartStampError))
                {
                    fails.Add("没写上对照数字：" + ChartStampError);
                }
                else if (page.IndexOf(ChartStamp, StringComparison.Ordinal) >= 0)
                {
                    fails.Add("图表数据 " + ChartStamp + " 进了概览");
                }

                Finish(run, "ovt-chart-title", fails);
            }

            if (Want(filters, "ovt-group-text"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 5);
                if (page.IndexOf("组内甲", StringComparison.Ordinal) < 0
                    || page.IndexOf("组内乙", StringComparison.Ordinal) < 0)
                {
                    fails.Add("组内文字没出来");
                }

                Finish(run, "ovt-group-text", fails);
            }

            if (Want(filters, "ovt-empty"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                string page = PageBody(display, 6).Trim();
                if (page != "（无文字）")
                {
                    fails.Add("空页不是「（无文字）」，实际 " + page);
                }

                Finish(run, "ovt-empty", fails);
            }

            if (Want(filters, "ovt-json-slim"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                if (slidesJson.IndexOf(LongSentence, StringComparison.Ordinal) >= 0
                    || slidesJson.IndexOf(ChartTitle, StringComparison.Ordinal) >= 0
                    || slidesJson.IndexOf("PageText", StringComparison.Ordinal) >= 0)
                {
                    fails.Add("slides JSON 抄了正文");
                }

                if (slidesJson.IndexOf("\"index\"", StringComparison.Ordinal) < 0
                    || slidesJson.IndexOf("slide_id", StringComparison.Ordinal) < 0)
                {
                    fails.Add("slides 缺 index / slide_id");
                }

                Finish(run, "ovt-json-slim", fails);
            }

            if (Want(filters, "ovt-slide-id-read"))
            {
                var fails = new List<string>();
                RequireOk(result, fails);
                Match id = Regex.Match(display ?? "", @"index=1 slide_id=([0-9]+)");
                if (!id.Success)
                {
                    fails.Add("第 1 页没有 slide_id");
                }
                else
                {
                    ToolResult pageRead = InvokeRead(channelId, id.Groups[1].Value);
                    if (pageRead == null || !pageRead.Success)
                    {
                        fails.Add(pageRead == null ? "读页不可调" : (pageRead.Error ?? "读页失败"));
                    }
                    else if ((GetDisplay(pageRead) ?? "").IndexOf("跨应用原生操控", StringComparison.Ordinal) < 0)
                    {
                        fails.Add("用概览 slide_id 读回的不是第 1 页");
                    }
                }

                Finish(run, "ovt-slide-id-read", fails);
            }
        }

        private static void StampChartData(object chart)
        {
            dynamic data = chart.GetType().InvokeMember(
                "ChartData",
                System.Reflection.BindingFlags.GetProperty,
                null,
                chart,
                null);
            data.Activate();
            dynamic sheet = data.Workbook.Worksheets[1];
            sheet.Cells[2, 2].Value2 = 987654;
        }

        private static ToolResult InvokeRead(string channelId, string slideId)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_ReadPptHtmlTool.Register(tools, null);
            return tools["F_read_ppt_html"](new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["channel_id"] = channelId,
                ["slide_id"] = slideId,
                ["content_only"] = true
            }).GetAwaiter().GetResult();
        }

        private static void RequireOk(ToolResult result, List<string> fails)
        {
            if (result == null)
            {
                fails.Add("工具不可调");
            }
            else if (!result.Success)
            {
                fails.Add(result.Error ?? "失败");
            }
        }

        private static ToolResult Invoke(string channelId)
        {
            var tools = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>(
                StringComparer.OrdinalIgnoreCase);
            F_GetPresentationContentTool.Register(tools, null);
            return tools["F_get_presentation_content"](new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["channel_id"] = channelId
            }).GetAwaiter().GetResult();
        }

        private static string GetDisplay(ToolResult result)
        {
            var data = result != null ? result.Data as Dictionary<string, object> : null;
            if (data == null || !data.TryGetValue("display_contents", out object value) || value == null)
            {
                return "";
            }

            return Convert.ToString(value) ?? "";
        }

        private static string SlidesJson(ToolResult result)
        {
            var data = result != null ? result.Data as Dictionary<string, object> : null;
            if (data == null || !data.TryGetValue("slides", out object slides) || slides == null)
            {
                return "";
            }

            return JsonConvert.SerializeObject(slides);
        }

        private static string PageBody(string display, int index)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            var matches = Regex.Matches(
                display,
                @"index=(\d+) slide_id=[^\r\n]*\r?\n",
                RegexOptions.CultureInvariant);
            for (int i = 0; i < matches.Count; i++)
            {
                if (matches[i].Groups[1].Value != index.ToString(CultureInfo.InvariantCulture))
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

        private static void ClearShapes(PowerPoint.Slide slide)
        {
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
        }

        private static void ClearWppShapes(object slide)
        {
            try
            {
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
            }
            catch (Exception)
            {
            }
        }

        private static void TrySetLayoutBlank(object slide)
        {
            try
            {
                WppCom.TrySetProperty(slide, "Layout", 12);
            }
            catch (Exception)
            {
            }
        }

        private static void SetWppCell(object tableShape, int row, int col, string text)
        {
            object table = WppCom.GetProperty(tableShape, "Table");
            object cell = WppCom.Invoke(table, "Cell", row, col);
            object frame = WppCom.GetProperty(WppCom.GetProperty(cell, "Shape"), "TextFrame");
            WppCom.TrySetProperty(WppCom.GetProperty(frame, "TextRange"), "Text", text);
        }

        private static void WritePng(string path)
        {
            using (var bmp = new Bitmap(64, 64))
            {
                bmp.SetPixel(0, 0, Color.Red);
                bmp.Save(path, ImageFormat.Png);
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
                if (string.Equals(f, "ovt", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(f, "ovt-", StringComparison.OrdinalIgnoreCase))
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
    }
}
