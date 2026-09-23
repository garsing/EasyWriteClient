using System;
using System.Collections.Generic;
using System.Globalization;
using WordAddIn1;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    internal sealed class ContentGroupSession
    {
        public bool Wpp { get; set; }
        public object Presentation { get; set; }
        public object Slide { get; set; }
        public string SlideId { get; set; }
        public string Rpc { get; set; }
        public ContentAssets Assets { get; set; }

        public bool TryApply(string html, bool allowCreate, out List<string> ids, out string error)
        {
            ids = new List<string>();
            if (!ContentHtml.TryParse(html, SlideId, out PptHtmlApplyPlan plan, out error))
            {
                return false;
            }

            plan.AllowCreate = allowCreate;
            PptHtmlApplyResult result;
            bool ok = Wpp
                ? PptHtmlWppApplier.TryApply(Presentation, plan, "grp-test-wpp", out result, out error)
                : PptHtmlPowerPointApplier.TryApply(
                    (PowerPoint.Presentation)Presentation, plan, "grp-test", out result, out error);
            if (!ok)
            {
                return false;
            }

            ids = CollectIds(result);
            return true;
        }

        public bool TryReadPage(out PptHtmlReadResult result, out string error)
        {
            // 对齐 F_read_ppt_html 默认：骨架，不是 full 导出。
            return TryRead(null, false, out result, out error);
        }

        public bool TryReadFullPage(out PptHtmlReadResult result, out string error)
        {
            return TryRead(null, true, out result, out error);
        }

        public bool TryReadShape(string shapeId, out PptHtmlReadResult result, out string error)
        {
            return TryRead(shapeId, false, out result, out error);
        }

        public bool TryGroup(string[] shapeIds, out PresentationManageShapeResult result, out string error)
        {
            return TryManage(new PresentationManageShapeRequest
            {
                Action = "group",
                SlideId = SlideId,
                ShapeIds = shapeIds
            }, Presentation, out result, out error);
        }

        public bool TryDuplicate(
            string groupId,
            double left,
            double top,
            out PresentationManageShapeResult result,
            out string error,
            string toSlideId = null,
            object destPresentation = null)
        {
            return TryManage(new PresentationManageShapeRequest
            {
                Action = "duplicate_group",
                SlideId = SlideId,
                ToSlideId = toSlideId,
                ShapeId = groupId,
                LeftPct = left,
                TopPct = top
            }, destPresentation ?? Presentation, out result, out error);
        }

        public bool TryAddBlankSlide(out string slideId, out string error)
        {
            slideId = null;
            error = null;
            try
            {
                if (Wpp)
                {
                    object slides = WppCom.GetProperty(Presentation, "Slides");
                    int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                    object slide;
                    try
                    {
                        slide = WppCom.Invoke(slides, "Add", count + 1, 12);
                    }
                    catch
                    {
                        slide = WppCom.Invoke(slides, "Add", count + 1, 1);
                    }

                    slideId = Convert.ToInt32(WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId"))
                        .ToString(CultureInfo.InvariantCulture);
                    return true;
                }

                PowerPoint.Presentation pres = (PowerPoint.Presentation)Presentation;
                PowerPoint.Slide slidePpt = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                slideId = slidePpt.SlideID.ToString(CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public bool TryOpenSecondDeck(out object dest, out string destSlideId, out string error)
        {
            dest = null;
            destSlideId = null;
            error = null;
            try
            {
                if (Wpp)
                {
                    object app = WppCom.GetProperty(Presentation, "Application");
                    object presentations = WppCom.GetProperty(app, "Presentations");
                    try
                    {
                        dest = WppCom.Invoke(presentations, "Add", true);
                    }
                    catch
                    {
                        dest = WppCom.Invoke(presentations, "Add");
                    }

                    if (dest == null)
                    {
                        error = "第二份 WPP 稿 Add 失败";
                        return false;
                    }

                    object slide = AddWppBlank(dest);
                    destSlideId = Convert.ToInt32(WppCom.GetProperty(slide, "SlideID") ?? WppCom.GetProperty(slide, "SlideId"))
                        .ToString(CultureInfo.InvariantCulture);
                    return true;
                }

                PowerPoint.Presentation src = (PowerPoint.Presentation)Presentation;
                PowerPoint.Presentation second = src.Application.Presentations.Add(Office.MsoTriState.msoTrue);
                dest = second;
                if (second.Slides.Count == 0)
                {
                    second.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
                }

                destSlideId = second.Slides[1].SlideID.ToString(CultureInfo.InvariantCulture);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        public void CloseDeckQuiet(object dest)
        {
            if (dest == null)
            {
                return;
            }

            try
            {
                if (Wpp)
                {
                    WppCom.TrySetProperty(dest, "Saved", true);
                    WppCom.Invoke(dest, "Close");
                }
                else
                {
                    PowerPoint.Presentation p = (PowerPoint.Presentation)dest;
                    p.Saved = Office.MsoTriState.msoTrue;
                    p.Close();
                }
            }
            catch
            {
            }
        }

        public static bool ToolRejects(string action, Dictionary<string, object> args, string expectContains)
        {
            bool ok = F_ManagePptShapeTool.TryRejectForeignArgs(action, args, out string error);
            return !ok
                && !string.IsNullOrEmpty(error)
                && error.IndexOf(expectContains, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static PptHtmlShapeNode FindGroup(PptHtmlReadResult result)
        {
            return FindGroupIn(result?.Shapes);
        }

        public static PptHtmlShapeNode FindChildGroup(PptHtmlShapeNode parent)
        {
            return FindGroupIn(parent?.Children);
        }

        public static PptHtmlShapeNode FindGroupIn(IList<PptHtmlShapeNode> nodes)
        {
            if (nodes == null)
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                if (IsGroup(nodes[i]))
                {
                    return nodes[i];
                }
            }

            return null;
        }

        public static int CountGroups(PptHtmlReadResult result)
        {
            int n = 0;
            if (result?.Shapes == null)
            {
                return 0;
            }

            for (int i = 0; i < result.Shapes.Count; i++)
            {
                if (IsGroup(result.Shapes[i]))
                {
                    n++;
                }
            }

            return n;
        }

        public static PptHtmlShapeNode FindByIdDeep(PptHtmlReadResult result, string shapeId)
        {
            return FindByIdDeep(result?.Shapes, shapeId);
        }

        public static PptHtmlShapeNode FindByIdDeep(IList<PptHtmlShapeNode> nodes, string shapeId)
        {
            if (nodes == null || string.IsNullOrEmpty(shapeId))
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (string.Equals(n.ShapeId, shapeId, StringComparison.Ordinal))
                {
                    return n;
                }

                PptHtmlShapeNode hit = FindByIdDeep(n.Children, shapeId);
                if (hit != null)
                {
                    return hit;
                }
            }

            return null;
        }

        public static PptHtmlShapeNode FindByTypeDeep(PptHtmlReadResult result, string shapeType)
        {
            return FindByTypeDeep(result?.Shapes, shapeType);
        }

        public static PptHtmlShapeNode FindByTypeDeep(IList<PptHtmlShapeNode> nodes, string shapeType)
        {
            if (nodes == null || string.IsNullOrEmpty(shapeType))
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (string.Equals(n.ShapeType, shapeType, StringComparison.OrdinalIgnoreCase))
                {
                    return n;
                }

                PptHtmlShapeNode hit = FindByTypeDeep(n.Children, shapeType);
                if (hit != null)
                {
                    return hit;
                }
            }

            return null;
        }

        public static PptHtmlShapeNode FindTextTop(PptHtmlReadResult result, string contains)
        {
            return FindTextIn(result?.Shapes, contains, false);
        }

        public static PptHtmlShapeNode FindText(PptHtmlReadResult result, string contains)
        {
            return FindTextIn(result?.Shapes, contains, true);
        }

        public static PptHtmlShapeNode FindTextIn(IList<PptHtmlShapeNode> nodes, string contains, bool deep)
        {
            if (nodes == null || string.IsNullOrEmpty(contains))
            {
                return null;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(n.Text)
                    && n.Text.IndexOf(contains, StringComparison.Ordinal) >= 0)
                {
                    return n;
                }

                if (deep)
                {
                    PptHtmlShapeNode hit = FindTextIn(n.Children, contains, true);
                    if (hit != null)
                    {
                        return hit;
                    }
                }
            }

            return null;
        }

        public static bool IsGroup(PptHtmlShapeNode n)
        {
            return n != null
                && string.Equals(n.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
        }

        public static bool HasChildNodes(PptHtmlShapeNode n)
        {
            return n != null && n.Children != null && n.Children.Count > 0;
        }

        public static bool GeoNear(string style, double left, double top, double eps = 1.5)
        {
            if (!ContentAssert.TryParseGeo(style, out double l, out double t, out _, out _))
            {
                return false;
            }

            return Math.Abs(l - left) <= eps && Math.Abs(t - top) <= eps;
        }

        public static bool TopNearGroup(PptHtmlShapeNode child, PptHtmlShapeNode group, double slop = 8)
        {
            if (child == null || group == null
                || !ContentAssert.TryParseGeo(child.Style, out _, out double ct, out _, out _)
                || !ContentAssert.TryParseGeo(group.Style, out _, out double gt, out _, out double gh))
            {
                return false;
            }

            if (Math.Abs(ct) < 0.8)
            {
                return false;
            }

            return ct + 0.5 >= gt - slop && ct <= gt + gh + slop;
        }

        public static List<PptHtmlShapeNode> DirectChildren(PptHtmlReadResult focused, string groupId)
        {
            PptHtmlShapeNode group = FindByIdDeep(focused, groupId) ?? FindGroup(focused);
            return group?.Children ?? new List<PptHtmlShapeNode>();
        }

        public static List<string> LeafIds(PptHtmlReadResult expand)
        {
            var ids = new List<string>();
            CollectLeafIds(expand?.Shapes, ids);
            return ids;
        }

        public static string DescribeTree(PptHtmlReadResult result)
        {
            var sb = new System.Text.StringBuilder();
            AppendTree(result?.Shapes, sb, 0);
            return sb.Length == 0 ? "(空)" : sb.ToString();
        }

        private static void AppendTree(IList<PptHtmlShapeNode> nodes, System.Text.StringBuilder sb, int depth)
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

                if (sb.Length > 0)
                {
                    sb.Append(" | ");
                }

                sb.Append(new string('.', depth))
                    .Append(n.ShapeType ?? "?")
                    .Append(':')
                    .Append(n.ShapeId ?? "-");
                if (!string.IsNullOrEmpty(n.Text))
                {
                    sb.Append('[').Append(n.Text.Replace("\n", " ").Trim()).Append(']');
                }

                AppendTree(n.Children, sb, depth + 1);
            }
        }

        private static void CollectLeafIds(IList<PptHtmlShapeNode> nodes, List<string> ids)
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

                if (!IsGroup(n) && !string.IsNullOrEmpty(n.ShapeId))
                {
                    ids.Add(n.ShapeId);
                }

                CollectLeafIds(n.Children, ids);
            }
        }

        private bool TryRead(string shapeId, bool fullPage, out PptHtmlReadResult result, out string error)
        {
            if (Wpp)
            {
                return PptHtmlWppReader.TryRead(
                    Presentation,
                    SlideId,
                    "grp-test-wpp",
                    "wpp",
                    shapeId,
                    fullPage,
                    false,
                    false,
                    false,
                    0,
                    out result,
                    out error);
            }

            return PptHtmlPowerPointReader.TryRead(
                (PowerPoint.Presentation)Presentation,
                SlideId,
                "grp-test",
                "ppt",
                shapeId,
                fullPage,
                false,
                false,
                false,
                0,
                out result,
                out error);
        }

        private bool TryManage(
            PresentationManageShapeRequest request,
            object dest,
            out PresentationManageShapeResult result,
            out string error)
        {
            if (Wpp)
            {
                return PptShapeWppManager.TryManage(
                    dest, Presentation, "wpp-test", "wpp-test", request, out result, out error);
            }

            return PptShapePowerPointManager.TryManage(
                (PowerPoint.Presentation)dest,
                (PowerPoint.Presentation)Presentation,
                "ppt-test",
                "ppt-test",
                request,
                out result,
                out error);
        }

        private static object AddWppBlank(object presentation)
        {
            object slides = WppCom.GetProperty(presentation, "Slides");
            int count = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
            if (count == 0)
            {
                try
                {
                    return WppCom.Invoke(slides, "Add", 1, 12);
                }
                catch
                {
                    return WppCom.Invoke(slides, "Add", 1, 1);
                }
            }

            return WppCom.GetIndexed(slides, 1);
        }

        private static List<string> CollectIds(PptHtmlApplyResult result)
        {
            var ids = new List<string>();
            if (result?.CreatedShapes == null)
            {
                return ids;
            }

            for (int i = 0; i < result.CreatedShapes.Count; i++)
            {
                Dictionary<string, object> d = result.CreatedShapes[i];
                if (d != null && d.TryGetValue("shape_id", out object v) && v != null)
                {
                    string id = v.ToString();
                    if (!string.IsNullOrEmpty(id))
                    {
                        ids.Add(id);
                    }
                }
            }

            return ids;
        }
    }

    internal static class ContentGroupHarness
    {
        public static string Run(
            TestRun run,
            object presentation,
            object slide,
            string slideId,
            ContentCase one,
            string tag,
            bool wpp,
            ContentAssets assets = null)
        {
            var session = new ContentGroupSession
            {
                Wpp = wpp,
                Presentation = presentation,
                Slide = slide,
                SlideId = slideId,
                Assets = assets
            };

            string mismatch;
            try
            {
                mismatch = one.GroupRun(session);
            }
            catch (Exception ex)
            {
                if (IsRpc(ex.Message))
                {
                    return ex.Message;
                }

                run.CaseFail(tag, "组脚本异常: " + ex.Message);
                return null;
            }

            if (!string.IsNullOrEmpty(session.Rpc) && IsRpc(session.Rpc))
            {
                return session.Rpc;
            }

            if (mismatch == null)
            {
                run.CaseOk(tag);
            }
            else if (mismatch.StartsWith("SKIP:", StringComparison.OrdinalIgnoreCase))
            {
                run.CaseSkip(tag, mismatch.Substring(5).Trim());
            }
            else
            {
                run.CaseFail(tag, mismatch);
            }

            return null;
        }

        private static bool IsRpc(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return false;
            }

            return message.IndexOf("RPC", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("0x800706BA", StringComparison.OrdinalIgnoreCase) >= 0
                || message.IndexOf("被调用的对象已与其客户端断开", StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
