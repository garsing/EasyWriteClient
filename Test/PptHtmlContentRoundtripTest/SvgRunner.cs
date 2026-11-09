using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>约定 HTML 的内嵌 svg 与 .svg 文件引用。前缀 svg-。</summary>
    internal static class SvgRunner
    {
        private const string HousePath = "M4 12 L12 4 L20 12 H17 V20 H7 V12 Z";

        private const string HouseSvg =
            "<svg viewBox=\"0 0 24 24\"><path d=\"" + HousePath + "\"/></svg>";

        private static readonly string[] HostIds =
        {
            "svg-inline-rt",
            "svg-inline-box",
            "svg-file-disk",
            "svg-file-rt",
            "svg-png-keep",
            "svg-script",
            "svg-foreign",
            "svg-event",
            "svg-href",
            "svg-scheme",
            "svg-big",
            "svg-update",
            "svg-skel",
            "svg-search"
        };

        public static bool Wants(IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.StartsWith("svg", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        public static void RunOffline(TestRun run, IList<string> filters)
        {
            Console.WriteLine("--- svg 解析与技能（不启宿主）---");
            if (Want(filters, "svg-parse-inline"))
            {
                Finish(run, "svg-parse-inline", ParseInline());
            }

            if (Want(filters, "svg-parse-script"))
            {
                Finish(run, "svg-parse-script", ParseReject(SvgWith("<script>bad</script>"), "script"));
            }

            if (Want(filters, "svg-parse-file"))
            {
                Finish(run, "svg-parse-file", ParseFile());
            }

            if (Want(filters, "svg-parse-empty"))
            {
                Finish(run, "svg-parse-empty", ParseEmpty());
            }

            if (Want(filters, "svg-serialize"))
            {
                Finish(run, "svg-serialize", Serialize());
            }

            if (Want(filters, "svg-skel-keep"))
            {
                Finish(run, "svg-skel-keep", SkelKeep());
            }

            if (Want(filters, "svg-skill-attrs"))
            {
                Finish(run, "svg-skill-attrs", SkillAttrs());
            }

            if (Want(filters, "svg-skill-diagram"))
            {
                Finish(run, "svg-skill-diagram", SkillDiagram());
            }
        }

        public static void RunHost(TestRun run, IList<string> filters, bool useWpp)
        {
            if (!NeedHost(filters))
            {
                return;
            }

            Console.WriteLine(useWpp ? "--- svg 宿主（WPP）---" : "--- svg 宿主（PPT）---");
            if (!UserService.Instance.EnsureEffectiveRootInteractive(null))
            {
                FailAll(run, filters, "工作区未就绪");
                return;
            }

            string exportDir = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                useWpp ? "svg-export-wpp" : "svg-export-ppt");
            Directory.CreateDirectory(exportDir);
            if (useWpp)
            {
                RunWpp(run, filters, exportDir);
            }
            else
            {
                RunPpt(run, filters, exportDir);
            }
        }

        private static void RunPpt(TestRun run, IList<string> filters, string exportDir)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            try
            {
                app = new PowerPoint.Application();
                app.Visible = Office.MsoTriState.msoTrue;
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch (Exception)
                {
                }

                pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                var session = new Session(pres, exportDir, wpp: false);
                RunPages(run, filters, session);
            }
            catch (Exception ex)
            {
                FailAll(run, filters, ex.Message);
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
            }
        }

        private static void RunWpp(TestRun run, IList<string> filters, string exportDir)
        {
            object app = null;
            object pres = null;
            try
            {
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    FailAll(run, filters, startErr);
                    return;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                try
                {
                    pres = WppCom.Invoke(presentations, "Add", true);
                }
                catch (Exception)
                {
                    pres = WppCom.Invoke(presentations, "Add");
                }

                if (pres == null)
                {
                    FailAll(run, filters, "WPP Presentations.Add 失败");
                    return;
                }

                var session = new Session(pres, exportDir, wpp: true);
                RunPages(run, filters, session);
            }
            catch (Exception ex)
            {
                FailAll(run, filters, ex.Message);
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
                    if (app != null)
                    {
                        object presentations = WppCom.GetProperty(app, "Presentations");
                        int n = Convert.ToInt32(WppCom.GetProperty(presentations, "Count") ?? 0);
                        if (n == 0)
                        {
                            WppCom.Invoke(app, "Quit");
                        }
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static void RunPages(TestRun run, IList<string> filters, Session session)
        {
            RunInline(run, filters, session);
            RunFile(run, filters, session);
            RunReject(run, filters, session);
        }

        private static void RunInline(TestRun run, IList<string> filters, Session session)
        {
            bool need = Want(filters, "svg-inline-rt")
                || Want(filters, "svg-inline-box")
                || Want(filters, "svg-file-disk")
                || Want(filters, "svg-update")
                || Want(filters, "svg-skel")
                || Want(filters, "svg-search");
            if (!need)
            {
                return;
            }

            if (!session.TryNewSlide(out string error))
            {
                FailThese(run, filters, error, "svg-inline-rt", "svg-inline-box", "svg-file-disk", "svg-update", "svg-skel", "svg-search");
                return;
            }

            if (!session.TryApply(InlineHtml(null, HousePath), out PptHtmlApplyResult created, out error))
            {
                FailThese(run, filters, error, "svg-inline-rt", "svg-inline-box", "svg-file-disk", "svg-update", "svg-skel", "svg-search");
                return;
            }

            if (!session.TryRead(contentOnly: false, out PptHtmlReadResult full, out error))
            {
                FailThese(run, filters, error, "svg-inline-rt", "svg-inline-box", "svg-file-disk", "svg-update", "svg-skel", "svg-search");
                return;
            }

            PptHtmlShapeNode svg = FindType(full, "svg");
            if (Want(filters, "svg-inline-rt"))
            {
                var fails = new List<string>();
                if (svg == null)
                {
                    fails.Add("读回没有 svg");
                }
                else
                {
                    if (string.IsNullOrEmpty(svg.ShapeId))
                    {
                        fails.Add("没有 ShapeId");
                    }

                    if (svg.SvgMarkup == null || svg.SvgMarkup.IndexOf(HousePath, StringComparison.Ordinal) < 0)
                    {
                        fails.Add("路径不在：" + (svg.SvgMarkup ?? ""));
                    }

                    if (!string.IsNullOrEmpty(svg.RasterizedFrom))
                    {
                        fails.Add("带了 data-rasterized-from=" + svg.RasterizedFrom);
                    }

                    string html = PptConventionHtml.BuildDisplayContents(One(svg));
                    if (html.IndexOf("<svg", StringComparison.Ordinal) < 0
                        || html.IndexOf("&lt;svg", StringComparison.Ordinal) >= 0)
                    {
                        fails.Add("序列化不是 svg 元素");
                    }
                }

                Finish(run, "svg-inline-rt", fails);
            }

            if (Want(filters, "svg-inline-box"))
            {
                var fails = new List<string>();
                string style = svg == null ? "" : (svg.Style ?? "");
                if (!Near(style, "left", 10) || !Near(style, "top", 20)
                    || !Near(style, "width", 15) || !Near(style, "height", 15))
                {
                    fails.Add("style=" + style);
                }

                Finish(run, "svg-inline-box", fails);
            }

            if (Want(filters, "svg-file-disk"))
            {
                Finish(run, "svg-file-disk", CheckDisk(svg));
            }

            if (Want(filters, "svg-skel"))
            {
                var fails = new List<string>();
                if (!session.TryRead(contentOnly: true, out PptHtmlReadResult skel, out error))
                {
                    fails.Add(error ?? "骨架读失败");
                }
                else
                {
                    PptHtmlShapeNode kept = FindType(skel, "svg");
                    if (kept == null)
                    {
                        fails.Add("骨架没有 svg");
                    }
                    else if (!string.IsNullOrEmpty(kept.RasterizedFrom))
                    {
                        fails.Add("骨架带了 data-rasterized-from");
                    }
                }

                Finish(run, "svg-skel", fails);
            }

            if (Want(filters, "svg-search"))
            {
                Finish(run, "svg-search", Search(full, svg == null ? "" : svg.ShapeId));
            }

            if (Want(filters, "svg-update"))
            {
                var fails = new List<string>();
                string oldId = svg == null ? "" : svg.ShapeId;
                if (string.IsNullOrEmpty(oldId))
                {
                    fails.Add("没有可更新的 ShapeId");
                }
                else if (!session.TryApply(InlineHtml(oldId, "M1 1"), out PptHtmlApplyResult updated, out error))
                {
                    fails.Add(error ?? "更新失败");
                }
                else
                {
                    string warn = JoinWarnings(updated);
                    if (warn.IndexOf("ShapeId 变为", StringComparison.Ordinal) >= 0)
                    {
                        fails.Add(warn);
                    }

                    if (!session.TryRead(contentOnly: false, out PptHtmlReadResult after, out error))
                    {
                        fails.Add(error ?? "更新后读失败");
                    }
                    else
                    {
                        PptHtmlShapeNode again = FindId(after, oldId);
                        if (again == null)
                        {
                            fails.Add("更新后找不到原 ShapeId " + oldId);
                        }
                        else
                        {
                            string markup = again.SvgMarkup ?? "";
                            if (markup.IndexOf("M1 1", StringComparison.Ordinal) < 0)
                            {
                                fails.Add("新路径不在：" + markup);
                            }

                            if (markup.IndexOf(HousePath, StringComparison.Ordinal) >= 0)
                            {
                                fails.Add("旧路径还在");
                            }
                        }
                    }
                }

                if (created == null)
                {
                }

                Finish(run, "svg-update", fails);
            }
        }

        private static void RunFile(TestRun run, IList<string> filters, Session session)
        {
            bool need = Want(filters, "svg-file-rt") || Want(filters, "svg-png-keep");
            if (!need)
            {
                return;
            }

            if (!session.TryNewSlide(out string error))
            {
                FailThese(run, filters, error, "svg-file-rt", "svg-png-keep");
                return;
            }

            string svgPath = WriteHouseFile();
            if (svgPath == null)
            {
                FailThese(run, filters, "写 icons/house.svg 失败", "svg-file-rt", "svg-png-keep");
                return;
            }

            string png = Path.Combine(session.ExportDir, "keep.png");
            File.WriteAllBytes(png, ContentAssetsIo.MinimalPngRed);
            string html = ContentHtml.Section(
                "  <img data-shape-type=\"picture\" data-src=\"icons/house.svg\" style=\""
                + ContentHtml.Geo(10, 20, 15, 15) + "\" />\n"
                + "  <img data-shape-type=\"picture\" data-src=\"" + png.Replace("&", "&amp;") + "\" style=\""
                + ContentHtml.Geo(40, 20, 15, 15) + "\" />");
            if (!session.TryApply(html, out _, out error)
                || !session.TryRead(contentOnly: false, out PptHtmlReadResult full, out error))
            {
                FailThese(run, filters, error, "svg-file-rt", "svg-png-keep");
                return;
            }

            string exportErr = ContentAssert.AttachPictureSrc(session.Pres, full, session.ExportDir);
            if (Want(filters, "svg-file-rt"))
            {
                var fails = new List<string>();
                PptHtmlShapeNode file = FindSrcEnding(full, ".svg");
                if (file == null || !string.Equals(file.ShapeType, "picture", StringComparison.OrdinalIgnoreCase))
                {
                    fails.Add("没有 picture data-src 以 .svg 结尾");
                }
                else if (file.DataSrc.IndexOf("icons/house.svg", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    fails.Add("data-src=" + file.DataSrc);
                }
                else if ((file.SvgMarkup ?? "").IndexOf("<path", StringComparison.OrdinalIgnoreCase) >= 0
                    && string.Equals(file.ShapeType, "svg", StringComparison.OrdinalIgnoreCase))
                {
                    fails.Add("文件引用被展开成 svg");
                }

                if (!string.IsNullOrEmpty(exportErr))
                {
                    fails.Add(exportErr);
                }

                Finish(run, "svg-file-rt", fails);
            }

            if (Want(filters, "svg-png-keep"))
            {
                var fails = new List<string>();
                PptHtmlShapeNode pic = FindSrcEnding(full, ".png");
                if (pic == null)
                {
                    fails.Add("没有导出 png");
                }
                else if (!File.Exists(pic.DataSrc) || !IsPng(pic.DataSrc))
                {
                    fails.Add("png 打不开：" + pic.DataSrc);
                }

                PptHtmlShapeNode file = FindSrcEnding(full, ".svg");
                if (file == null || file.DataSrc.IndexOf(".svg", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    fails.Add("svg 引用被带成别的后缀");
                }

                Finish(run, "svg-png-keep", fails);
            }
        }

        private static void RunReject(TestRun run, IList<string> filters, Session session)
        {
            bool need = Want(filters, "svg-script")
                || Want(filters, "svg-foreign")
                || Want(filters, "svg-event")
                || Want(filters, "svg-href")
                || Want(filters, "svg-scheme")
                || Want(filters, "svg-big");
            if (!need)
            {
                return;
            }

            if (!session.TryNewSlide(out string error))
            {
                FailThese(run, filters, error, "svg-script", "svg-foreign", "svg-event", "svg-href", "svg-scheme", "svg-big");
                return;
            }

            Reject(run, filters, session, "svg-script", SvgWith("<script>bad</script>"), "script");
            Reject(run, filters, session, "svg-foreign", SvgWith("<foreignObject></foreignObject>"), "foreign");
            Reject(run, filters, session, "svg-event", SvgWith("", " onclick=\"x\""), "事件");
            Reject(run, filters, session, "svg-href", SvgWith("<a href=\"https://example.com/a.svg\"></a>"), "链接");
            if (Want(filters, "svg-scheme"))
            {
                var fails = new List<string>();
                fails.AddRange(RejectOne(session, PictureSrc("https://example.com/a.svg"), "http"));
                fails.AddRange(RejectOne(session, PictureSrc("data:image/svg+xml;base64,AAAA"), "data"));
                Finish(run, "svg-scheme", fails);
            }

            if (Want(filters, "svg-big"))
            {
                string huge = new string('A', (512 * 1024) + 64);
                Reject(run, filters, session, "svg-big", SvgWith("<path d=\"" + huge + "\"/>"), "过大");
            }
        }

        private static void Reject(TestRun run, IList<string> filters, Session session, string id, string html, string needle)
        {
            if (!Want(filters, id))
            {
                return;
            }

            Finish(run, id, RejectOne(session, html, needle));
        }

        private static List<string> RejectOne(Session session, string html, string needle)
        {
            var fails = new List<string>();
            int before = session.ShapeCount();
            bool ok = session.TryApply(html, out _, out string error);
            int after = session.ShapeCount();
            if (ok)
            {
                fails.Add("应失败却成功");
            }
            else if (error == null || error.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("error=" + error);
            }

            if (before != after)
            {
                fails.Add("形状数 " + before + " → " + after);
            }

            return fails;
        }

        private static List<string> ParseInline()
        {
            var fails = new List<string>();
            if (!ContentHtml.TryParse(InlineHtml(null, HousePath), "260", out PptHtmlApplyPlan plan, out string error))
            {
                fails.Add(error ?? "解析失败");
                return fails;
            }

            PptHtmlApplyNode n = plan.Nodes[0];
            if (!n.IsCreate || n.ShapeType != "svg" || n.SvgMode != "inline")
            {
                fails.Add("type=" + n.ShapeType + " mode=" + n.SvgMode + " create=" + n.IsCreate);
            }

            if (n.SvgMarkup == null || n.SvgMarkup.IndexOf(HousePath, StringComparison.Ordinal) < 0)
            {
                fails.Add("原文不含路径");
            }

            if (n.HasText || !string.IsNullOrEmpty(n.Text))
            {
                fails.Add("正文=" + n.Text);
            }

            return fails;
        }

        private static List<string> ParseReject(string html, string needle)
        {
            var fails = new List<string>();
            bool ok = ContentHtml.TryParse(html, "260", out _, out string error);
            if (ok)
            {
                fails.Add("应失败却成功");
            }
            else if (error == null || error.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
            {
                fails.Add("error=" + error);
            }

            return fails;
        }

        private static List<string> ParseFile()
        {
            var fails = new List<string>();
            string html = ContentHtml.Section(
                "  <img data-shape-type=\"picture\" data-src=\"icons/house.svg\" style=\""
                + ContentHtml.Geo(10, 20, 15, 15) + "\" />");
            if (!ContentHtml.TryParse(html, "260", out PptHtmlApplyPlan plan, out string error))
            {
                fails.Add(error ?? "解析失败");
                return fails;
            }

            PptHtmlApplyNode n = plan.Nodes[0];
            if (n.ShapeType != "picture" || n.DataSrc != "icons/house.svg" || !string.IsNullOrEmpty(n.SvgMode))
            {
                fails.Add("type=" + n.ShapeType + " src=" + n.DataSrc + " mode=" + n.SvgMode);
            }

            return fails;
        }

        private static List<string> ParseEmpty()
        {
            var fails = new List<string>();
                fails.AddRange(ParseReject(SvgWith(""), "svg"));
            fails.AddRange(ParseReject(SvgWith("   "), "svg"));
            return fails;
        }

        private static List<string> Serialize()
        {
            var fails = new List<string>();
            string html = PptConventionHtml.BuildDisplayContents(One(new PptHtmlShapeNode
            {
                ShapeId = "sid260-s3",
                ShapeType = "svg",
                Tag = "svg",
                Style = ContentHtml.Geo(10, 20, 15, 15),
                SvgMode = "inline",
                SvgMarkup = HouseSvg,
                Editable = true
            }));
            if (html.IndexOf("<svg", StringComparison.Ordinal) < 0
                || html.IndexOf("sid260-s3", StringComparison.Ordinal) < 0
                || html.IndexOf(HousePath, StringComparison.Ordinal) < 0
                || html.IndexOf("&lt;svg", StringComparison.Ordinal) >= 0)
            {
                fails.Add(html);
            }

            return fails;
        }

        private static List<string> SkelKeep()
        {
            var fails = new List<string>();
            var shapes = new List<PptHtmlShapeNode>
            {
                new PptHtmlShapeNode
                {
                    ShapeType = "svg",
                    Tag = "svg",
                    Text = "",
                    SvgMarkup = HouseSvg
                }
            };
            PptConventionHtml.ApplyContentOnlyFilter(shapes, 1);
            PptConventionHtml.ApplyCompactFilter(shapes);
            if (shapes.Count != 1 || !string.IsNullOrEmpty(shapes[0].RasterizedFrom))
            {
                fails.Add("count=" + shapes.Count);
            }

            return fails;
        }

        private static List<string> SkillAttrs()
        {
            var fails = new List<string>();
            string path = FindRepoFile(@"WordAddInBackend\main-server\project\AgentSkills\skills\general\ppt-html\references\shape-attrs.md");
            if (path == null)
            {
                fails.Add("找不到 shape-attrs.md");
                return fails;
            }

            string text = File.ReadAllText(path, Encoding.UTF8);
            if (text.IndexOf("`svg`", StringComparison.Ordinal) < 0)
            {
                fails.Add("允许标签没有 svg");
            }

            if (text.IndexOf(".svg", StringComparison.Ordinal) < 0)
            {
                fails.Add("没有 .svg 文件引用");
            }

            return fails;
        }

        private static List<string> SkillDiagram()
        {
            var fails = new List<string>();
            string path = FindRepoFile(@"WordAddInBackend\main-server\project\AgentSkills\skills\general\create_diagram\SKILL.md");
            if (path == null)
            {
                fails.Add("找不到 create_diagram/SKILL.md");
                return fails;
            }

            string text = File.ReadAllText(path, Encoding.UTF8);
            if (text.IndexOf("F_render_svg", StringComparison.Ordinal) < 0
                || text.IndexOf("PNG", StringComparison.Ordinal) < 0)
            {
                fails.Add("没有先渲 PNG");
            }

            if (text.IndexOf("不要把 SVG 直接塞进插图工具", StringComparison.Ordinal) < 0)
            {
                fails.Add("没有禁止把 SVG 直接塞进插图工具");
            }

            if (text.IndexOf("data-shape-type=\"svg\"", StringComparison.Ordinal) >= 0)
            {
                fails.Add("技能改成了直接写约定 HTML 的 svg");
            }

            return fails;
        }

        private static List<string> CheckDisk(PptHtmlShapeNode svg)
        {
            var fails = new List<string>();
            if (svg == null || !PptShapeId.TryParseShape(svg.ShapeId, out string slideId, out int comId))
            {
                fails.Add("ShapeId 无法解析");
                return fails;
            }

            string file = Path.Combine(
                WorkspacePathResolver.GetSessionDirectory(),
                "ppt_svg",
                slideId + "_" + comId.ToString(CultureInfo.InvariantCulture) + ".svg");
            if (!File.Exists(file))
            {
                fails.Add("没有文件 " + file);
                return fails;
            }

            string text = File.ReadAllText(file, Encoding.UTF8);
            if (text.IndexOf(HousePath, StringComparison.Ordinal) < 0)
            {
                fails.Add("文件里没有原路径");
            }

            return fails;
        }

        private static List<string> Search(PptHtmlReadResult full, string svgId)
        {
            var fails = new List<string>();
            if (!MatchCount(full, "text", "M4 12", out int textHits, out string error))
            {
                fails.Add(error ?? "路径搜索失败");
            }
            else if (textHits != 0)
            {
                fails.Add("路径命中 " + textHits);
            }

            if (!MatchCount(full, "data-shape-type", "^svg$", out int typeHits, out List<string> ids, out error))
            {
                fails.Add(error ?? "类型搜索失败");
            }
            else if (typeHits != 1 || ids.Count != 1 || ids[0] != svgId)
            {
                fails.Add("类型命中 " + typeHits + " ids=" + string.Join(",", ids));
            }

            return fails;
        }

        private static bool MatchCount(PptHtmlReadResult full, string attr, string pattern, out int count, out string error)
        {
            return MatchCount(full, attr, pattern, out count, out _, out error);
        }

        private static bool MatchCount(
            PptHtmlReadResult full,
            string attr,
            string pattern,
            out int count,
            out List<string> ids,
            out string error)
        {
            count = 0;
            ids = new List<string>();
            var args = new Dictionary<string, object>
            {
                ["attr"] = attr,
                ["pattern"] = pattern
            };
            if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest request, out error))
            {
                return false;
            }

            if (!PptHtmlNodeSearch.TryMatch(
                    full.Shapes,
                    request,
                    out List<PptHtmlShapeNode> forest,
                    out int leafCount,
                    out _,
                    out _,
                    out error))
            {
                return false;
            }

            count = leafCount;
            CollectIds(forest, ids);
            return true;
        }

        private static void CollectIds(IList<PptHtmlShapeNode> nodes, List<string> ids)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (n.Children != null && n.Children.Count > 0)
                {
                    CollectIds(n.Children, ids);
                }
                else if (!string.IsNullOrEmpty(n.ShapeId))
                {
                    ids.Add(n.ShapeId);
                }
            }
        }

        private static string WriteHouseFile()
        {
            try
            {
                string dir = Path.Combine(WorkspacePathResolver.GetSessionDirectory(), "icons");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, "house.svg");
                File.WriteAllText(path, HouseSvg, new UTF8Encoding(false));
                return path;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string InlineHtml(string shapeId, string pathData)
        {
            string id = string.IsNullOrEmpty(shapeId) ? "" : " ShapeId=\"" + shapeId + "\"";
            return ContentHtml.Section(
                "  <svg" + id + " style=\"" + ContentHtml.Geo(10, 20, 15, 15)
                + "\" viewBox=\"0 0 24 24\"><path d=\"" + pathData + "\"/></svg>");
        }

        private static string SvgWith(string inner, string extraAttrs = null)
        {
            return ContentHtml.Section(
                "  <svg style=\"" + ContentHtml.Geo(10, 20, 15, 15) + "\""
                + (extraAttrs ?? "") + ">" + inner + "</svg>");
        }

        private static string PictureSrc(string src)
        {
            return ContentHtml.Section(
                "  <img data-shape-type=\"picture\" data-src=\"" + src + "\" style=\""
                + ContentHtml.Geo(10, 20, 15, 15) + "\" />");
        }

        private static PptHtmlReadResult One(PptHtmlShapeNode node)
        {
            return new PptHtmlReadResult
            {
                SlideId = "260",
                Shapes = new List<PptHtmlShapeNode> { node }
            };
        }

        private static PptHtmlShapeNode FindType(PptHtmlReadResult result, string type)
        {
            if (result?.Shapes == null)
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n != null && string.Equals(n.ShapeType, type, StringComparison.OrdinalIgnoreCase))
                {
                    return n;
                }
            }

            return null;
        }

        private static PptHtmlShapeNode FindId(PptHtmlReadResult result, string id)
        {
            if (result?.Shapes == null)
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n != null && string.Equals(n.ShapeId, id, StringComparison.Ordinal))
                {
                    return n;
                }
            }

            return null;
        }

        private static PptHtmlShapeNode FindSrcEnding(PptHtmlReadResult result, string suffix)
        {
            if (result?.Shapes == null)
            {
                return null;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                PptHtmlShapeNode n = result.Shapes[i];
                if (n != null
                    && n.DataSrc != null
                    && n.DataSrc.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                {
                    return n;
                }
            }

            return null;
        }

        private static bool Near(string style, string key, double want)
        {
            Match m = Regex.Match(
                style ?? "",
                key + @"\s*:\s*(-?\d+(?:\.\d+)?)%",
                RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!m.Success)
            {
                return false;
            }

            if (!double.TryParse(m.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double got))
            {
                return false;
            }

            return Math.Abs(got - want) <= 1.0;
        }

        private static bool IsPng(string path)
        {
            try
            {
                byte[] head = new byte[4];
                using (FileStream fs = File.OpenRead(path))
                {
                    if (fs.Read(head, 0, 4) < 4)
                    {
                        return false;
                    }
                }

                return head[0] == 0x89 && head[1] == 0x50 && head[2] == 0x4E && head[3] == 0x47;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string JoinWarnings(PptHtmlApplyResult result)
        {
            if (result?.Warnings == null || result.Warnings.Count == 0)
            {
                return "";
            }

            return string.Join(" ", result.Warnings);
        }

        private static string FindRepoFile(string relative)
        {
            DirectoryInfo dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null)
            {
                string candidate = Path.Combine(dir.FullName, relative);
                if (File.Exists(candidate))
                {
                    return candidate;
                }

                dir = dir.Parent;
            }

            return null;
        }

        private static bool NeedHost(IList<string> filters)
        {
            for (int i = 0; i < HostIds.Length; i++)
            {
                if (Want(filters, HostIds[i]))
                {
                    return true;
                }
            }

            return false;
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
                if (f.Length > 0 && id.IndexOf(f, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void Finish(TestRun run, string id, List<string> fails)
        {
            if (fails == null || fails.Count == 0)
            {
                run.CaseOk(id);
                return;
            }

            run.CaseFail(id, string.Join("；", fails));
        }

        private static void FailAll(TestRun run, IList<string> filters, string reason)
        {
            FailThese(run, filters, reason, HostIds);
        }

        private static void FailThese(TestRun run, IList<string> filters, string reason, params string[] ids)
        {
            for (int i = 0; i < ids.Length; i++)
            {
                if (Want(filters, ids[i]))
                {
                    run.CaseFail(ids[i], reason);
                }
            }
        }

        private sealed class Session
        {
            public Session(object pres, string exportDir, bool wpp)
            {
                Pres = pres;
                ExportDir = exportDir;
                Wpp = wpp;
            }

            public object Pres { get; private set; }

            public string ExportDir { get; private set; }

            private bool Wpp { get; set; }

            private object Slide { get; set; }

            private string SlideId { get; set; }

            public bool TryNewSlide(out string error)
            {
                error = null;
                try
                {
                    if (Wpp)
                    {
                        object slides = WppCom.GetProperty(Pres, "Slides");
                        int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                        try
                        {
                            Slide = WppCom.Invoke(slides, "Add", count + 1, 12);
                        }
                        catch (Exception)
                        {
                            Slide = WppCom.Invoke(slides, "Add", count + 1, 1);
                        }

                        SlideId = Convert.ToInt32(
                                WppCom.GetProperty(Slide, "SlideID") ?? WppCom.GetProperty(Slide, "SlideId"))
                            .ToString(CultureInfo.InvariantCulture);
                        return true;
                    }

                    PowerPoint.Presentation ppt = (PowerPoint.Presentation)Pres;
                    PowerPoint.Slide slide = ppt.Slides.Add(
                        ppt.Slides.Count + 1,
                        PowerPoint.PpSlideLayout.ppLayoutBlank);
                    Slide = slide;
                    SlideId = slide.SlideID.ToString(CultureInfo.InvariantCulture);
                    return true;
                }
                catch (Exception ex)
                {
                    error = ex.Message;
                    return false;
                }
            }

            public int ShapeCount()
            {
                try
                {
                    if (Wpp)
                    {
                        object shapes = WppCom.GetProperty(Slide, "Shapes");
                        return Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                    }

                    return ((PowerPoint.Slide)Slide).Shapes.Count;
                }
                catch (Exception)
                {
                    return -1;
                }
            }

            public bool TryApply(string html, out PptHtmlApplyResult result, out string error)
            {
                result = null;
                if (!ContentHtml.TryParse(html, SlideId, out PptHtmlApplyPlan plan, out error))
                {
                    return false;
                }

                plan.AllowCreate = true;
                error = AcceptFiles(plan);
                if (error != null)
                {
                    return false;
                }

                return Wpp
                    ? PptHtmlWppApplier.TryApply(Pres, plan, "svg-wpp", out result, out error)
                    : PptHtmlPowerPointApplier.TryApply(
                        (PowerPoint.Presentation)Pres,
                        plan,
                        "svg-ppt",
                        out result,
                        out error);
            }

            public bool TryRead(bool contentOnly, out PptHtmlReadResult result, out string error)
            {
                if (Wpp)
                {
                    return PptHtmlWppReader.TryRead(
                        Pres,
                        SlideId,
                        "svg-wpp",
                        "wpp",
                        null,
                        true,
                        false,
                        false,
                        contentOnly,
                        0,
                        out result,
                        out error);
                }

                return PptHtmlPowerPointReader.TryRead(
                    (PowerPoint.Presentation)Pres,
                    SlideId,
                    "svg-ppt",
                    "ppt",
                    null,
                    true,
                    false,
                    false,
                    contentOnly,
                    0,
                    out result,
                    out error);
            }

            private static string AcceptFiles(PptHtmlApplyPlan plan)
            {
                if (plan?.Nodes == null)
                {
                    return null;
                }

                for (int i = 0; i < plan.Nodes.Count; i++)
                {
                    string error = AcceptNode(plan.Nodes[i]);
                    if (error != null)
                    {
                        return error;
                    }
                }

                return null;
            }

            private static string AcceptNode(PptHtmlApplyNode node)
            {
                if (node == null || string.IsNullOrWhiteSpace(node.DataSrc))
                {
                    return null;
                }

                if (!PptHtmlFileSource.TryClassify(
                        node.DataSrc,
                        out bool absolute,
                        out bool workspace,
                        out string name,
                        out string error))
                {
                    return error;
                }

                if (absolute)
                {
                    string path = Path.GetFullPath(node.DataSrc.Trim());
                    if (!File.Exists(path))
                    {
                        return "文件不存在: " + path;
                    }

                    node.ResolvedLocalPath = path;
                }
                else if (workspace)
                {
                    string path = Path.Combine(
                        WorkspacePathResolver.GetSessionDirectory(),
                        name.Replace('/', Path.DirectorySeparatorChar));
                    if (!File.Exists(path))
                    {
                        return "文件不存在: " + name;
                    }

                    node.ResolvedLocalPath = path;
                }

                return PptHtmlSvgStore.TryAcceptFile(node);
            }
        }
    }
}
