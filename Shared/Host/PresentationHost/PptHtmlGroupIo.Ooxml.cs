using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    internal static partial class PptHtmlGroupIo
    {
        private static readonly XNamespace ANs =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

        private static readonly XNamespace PNs =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

        private static readonly XNamespace RNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static readonly XNamespace RelPkgNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private const double EmuPerPoint = 12700.0;

        private static bool TryGroupNestedByOoxmlPpt(
            PowerPoint.Slide slide,
            IList<PowerPoint.Shape> members,
            out PowerPoint.Shape group,
            out string error)
        {
            group = null;
            error = null;
            PowerPoint.Presentation pres;
            PowerPoint.Application app;
            try
            {
                pres = slide.Parent as PowerPoint.Presentation;
                app = pres == null ? null : pres.Application;
            }
            catch (Exception)
            {
                pres = null;
                app = null;
            }

            if (pres == null || app == null)
            {
                error = "OOXML 编组：没有 Presentation";
                return false;
            }

            var ids = new List<int>();
            if (!TryReadPptBox(members, ids, out float left, out float top, out float width, out float height, out error))
            {
                return false;
            }

            int slideIndex;
            try
            {
                slideIndex = slide.SlideIndex;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组：读不到 SlideIndex: " + ex.Message;
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "ew-ppt-group-" + Guid.NewGuid().ToString("N") + ".pptx");
            PowerPoint.Presentation copy = null;
            PowerPoint.PpAlertLevel prevAlerts = PowerPoint.PpAlertLevel.ppAlertsAll;
            bool alertsSet = false;
            try
            {
                try
                {
                    prevAlerts = app.DisplayAlerts;
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                    alertsSet = true;
                }
                catch (Exception)
                {
                }

                pres.SaveCopyAs(tempPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                if (!File.Exists(tempPath))
                {
                    error = "OOXML 编组：SaveCopyAs 未落盘";
                    return false;
                }

                if (!TryWrapMembersInPackage(
                    tempPath,
                    slideIndex,
                    ids,
                    left,
                    top,
                    width,
                    height,
                    out int newGroupId,
                    out string wrappedKinds,
                    out error))
                {
                    return false;
                }

                copy = TryOpenCopyPpt(app, tempPath, out error);
                if (copy == null)
                {
                    return false;
                }

                PowerPoint.Shape source = FindTopLevelPpt(copy.Slides[slideIndex].Shapes, newGroupId);
                if (source == null)
                {
                    error = "OOXML 编组：副本里找不到新组 id=" + newGroupId + " wrapped=" + wrappedKinds;
                    return false;
                }

                string probePath = tempPath + ".after-open.pptx";
                try
                {
                    copy.SaveCopyAs(probePath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                }
                catch (Exception)
                {
                    probePath = null;
                }

                bool xmlAlive = !string.IsNullOrEmpty(probePath)
                    && XmlGroupHasChildGroup(probePath, slideIndex, newGroupId);
                TryDeleteFile(probePath);
                if (!xmlAlive && !NestLooksAlivePpt(source))
                {
                    error = "OOXML 编组：打开副本后 XML/COM 都不见内组 wrapped=" + wrappedKinds
                        + " source=" + DescribePptGroupItems(source)
                        + " parents=" + DescribePptParents(source);
                    return false;
                }

                source.Copy();
                PowerPoint.ShapeRange pasted = slide.Shapes.Paste();
                if (pasted == null || pasted.Count < 1)
                {
                    error = "OOXML 编组：贴回失败";
                    return false;
                }

                PowerPoint.Shape created = pasted[1];
                try
                {
                    created.Left = left;
                    created.Top = top;
                }
                catch (Exception)
                {
                }

                if (!NestLooksAlivePpt(created) && !xmlAlive)
                {
                    error = "OOXML 编组：贴回后内组仍不在 wrapped=" + wrappedKinds
                        + " source=" + DescribePptGroupItems(source)
                        + " pasted=" + DescribePptGroupItems(created);
                    TryDeletePpt(created);
                    return false;
                }

                for (int i = 0; i < members.Count; i++)
                {
                    TryDeletePpt(members[i]);
                }

                PptHtmlGroupIo.InvalidateGroupTreeCache();
                group = created;
                return true;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        copy.Saved = Office.MsoTriState.msoTrue;
                        copy.Close();
                    }
                    catch (Exception)
                    {
                    }
                }

                if (alertsSet)
                {
                    try
                    {
                        app.DisplayAlerts = prevAlerts;
                    }
                    catch (Exception)
                    {
                    }
                }

                TryDeleteFile(tempPath);
            }
        }

        private static bool TryGroupNestedByOoxmlWpp(
            object slide,
            IList<object> members,
            out object group,
            out string error)
        {
            group = null;
            error = null;
            object pres = WppCom.GetProperty(slide, "Parent");
            object app = pres == null ? null : WppCom.GetProperty(pres, "Application");
            if (pres == null || app == null)
            {
                error = "OOXML 编组：没有 Presentation";
                return false;
            }

            var ids = new List<int>();
            if (!TryReadWppBox(members, ids, out double left, out double top, out double width, out double height, out error))
            {
                return false;
            }

            int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
            if (slideIndex < 1)
            {
                error = "OOXML 编组：SlideIndex 无效";
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "ew-wpp-group-" + Guid.NewGuid().ToString("N") + ".pptx");
            object copy = null;
            object prevAlerts = null;
            bool alertsSet = false;
            try
            {
                alertsSet = TrySilenceWppAlerts(app, out prevAlerts);
                if (!TrySaveCopyAsWpp(pres, tempPath, out error))
                {
                    return false;
                }

                if (!TryWrapMembersInPackage(
                    tempPath,
                    slideIndex,
                    ids,
                    left,
                    top,
                    width,
                    height,
                    out int newGroupId,
                    out string wrappedKinds,
                    out error))
                {
                    return false;
                }

                bool xmlAlive = wrappedKinds.IndexOf("grpSp", StringComparison.OrdinalIgnoreCase) >= 0;
                object presentations = WppCom.GetProperty(app, "Presentations");
                copy = TryOpenCopyWpp(presentations, tempPath, out error);
                if (copy == null)
                {
                    return false;
                }

                object copySlides = WppCom.GetProperty(copy, "Slides");
                object copySlide = WppCom.GetIndexed(copySlides, slideIndex);
                object copyShapes = copySlide == null ? null : WppCom.GetProperty(copySlide, "Shapes");
                object source = FindTopLevelWpp(copyShapes, newGroupId);
                if (source == null)
                {
                    error = "OOXML 编组：副本里找不到新组 id=" + newGroupId + " wrapped=" + wrappedKinds;
                    return false;
                }

                if (!xmlAlive && !NestLooksAliveWpp(source))
                {
                    error = "OOXML 编组：打开副本后 XML/COM 都不见内组 wrapped=" + wrappedKinds
                        + " source=" + DescribeWppGroupItems(source);
                    return false;
                }

                WppCom.Invoke(source, "Copy");
                object shapes = WppCom.GetProperty(slide, "Shapes");
                object pasted = TryPasteWpp(shapes);
                if (pasted == null)
                {
                    error = "OOXML 编组：贴回失败";
                    return false;
                }

                try
                {
                    WppCom.TrySetProperty(pasted, "Left", left);
                    WppCom.TrySetProperty(pasted, "Top", top);
                }
                catch (Exception)
                {
                }

                if (!NestLooksAliveWpp(pasted) && !xmlAlive)
                {
                    error = "OOXML 编组：贴回后内组仍不在 wrapped=" + wrappedKinds
                        + " pasted=" + DescribeWppGroupItems(pasted);
                    TryDeleteWpp(pasted);
                    return false;
                }

                for (int i = 0; i < members.Count; i++)
                {
                    TryDeleteWpp(members[i]);
                }

                InvalidateGroupTreeCache();
                group = pasted;
                return true;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        WppCom.Invoke(copy, "Close");
                    }
                    catch (Exception)
                    {
                    }
                }

                if (alertsSet)
                {
                    TryRestoreWppAlerts(app, prevAlerts);
                }

                TryDeleteFile(tempPath);
            }
        }

        private static bool TryWrapMembersInPackage(
            string pptxPath,
            int slideIndex,
            IList<int> memberIds,
            double left,
            double top,
            double width,
            double height,
            out int newGroupId,
            out string wrappedKinds,
            out string error)
        {
            newGroupId = 0;
            wrappedKinds = "";
            error = null;
            try
            {
                using (ZipArchive zip = ZipFile.Open(pptxPath, ZipArchiveMode.Update))
                {
                    List<string> slides = ListSlidePartPaths(zip);
                    if (slideIndex < 1 || slideIndex > slides.Count)
                    {
                        error = "OOXML 编组：slideIndex=" + slideIndex + " 超出 " + slides.Count;
                        return false;
                    }

                    ZipArchiveEntry entry = FindZipEntry(zip, slides[slideIndex - 1]);
                    if (entry == null)
                    {
                        error = "OOXML 编组：找不到 " + slides[slideIndex - 1];
                        return false;
                    }

                    XDocument doc;
                    using (Stream stream = entry.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    if (!TryWrapMembersInSlideXml(
                        doc,
                        memberIds,
                        left,
                        top,
                        width,
                        height,
                        out newGroupId,
                        out wrappedKinds,
                        out error))
                    {
                        return false;
                    }

                    string fullName = entry.FullName;
                    entry.Delete();
                    ZipArchiveEntry fresh = zip.CreateEntry(fullName, CompressionLevel.Optimal);
                    using (Stream stream = fresh.Open())
                    {
                        doc.Save(stream);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组改包失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryWrapMembersInSlideXml(
            XDocument doc,
            IList<int> memberIds,
            double left,
            double top,
            double width,
            double height,
            out int newGroupId,
            out string wrappedKinds,
            out string error)
        {
            newGroupId = 0;
            wrappedKinds = "";
            error = null;
            if (doc == null || doc.Root == null)
            {
                error = "OOXML 编组：slide XML 空";
                return false;
            }

            XElement spTree = null;
            foreach (XElement el in doc.Descendants(PNs + "spTree"))
            {
                spTree = el;
                break;
            }

            if (spTree == null)
            {
                error = "OOXML 编组：没有 p:spTree";
                return false;
            }

            var wanted = new Dictionary<int, bool>();
            for (int i = 0; i < memberIds.Count; i++)
            {
                wanted[memberIds[i]] = true;
            }

            var memberEls = new List<XElement>();
            foreach (XElement child in spTree.Elements())
            {
                string local = child.Name.LocalName;
                if (local == "nvGrpSpPr" || local == "grpSpPr")
                {
                    continue;
                }

                int id = ReadOwnShapeId(child);
                if (wanted.ContainsKey(id))
                {
                    memberEls.Add(child);
                    wanted.Remove(id);
                }
            }

            if (wanted.Count > 0 || memberEls.Count != memberIds.Count)
            {
                error = "OOXML 编组：页顶层找不到全部成员 found=" + memberEls.Count
                    + " expect=" + memberIds.Count;
                return false;
            }

            var kinds = new List<string>();
            for (int i = 0; i < memberEls.Count; i++)
            {
                kinds.Add(memberEls[i].Name.LocalName + ":" + ReadOwnShapeId(memberEls[i]));
            }

            wrappedKinds = string.Join(",", kinds.ToArray());
            newGroupId = MaxCnvPrId(doc.Root) + 1;
            long originX = ToEmuLong(left);
            long originY = ToEmuLong(top);
            XElement grp = BuildOuterGroup(newGroupId, left, top, width, height);
            memberEls[0].AddBeforeSelf(grp);
            for (int i = 0; i < memberEls.Count; i++)
            {
                ShiftAllXfrmInTree(memberEls[i], originX, originY);
                memberEls[i].Remove();
                grp.Add(memberEls[i]);
            }

            return true;
        }

        private static XElement BuildOuterGroup(
            int id,
            double left,
            double top,
            double width,
            double height)
        {
            string offX = ToEmu(left);
            string offY = ToEmu(top);
            string extCx = ToEmu(Math.Max(1.0, width));
            string extCy = ToEmu(Math.Max(1.0, height));
            return new XElement(
                PNs + "grpSp",
                new XElement(
                    PNs + "nvGrpSpPr",
                    new XElement(
                        PNs + "cNvPr",
                        new XAttribute("id", id.ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("name", "Group " + id.ToString(CultureInfo.InvariantCulture))),
                    new XElement(PNs + "cNvGrpSpPr"),
                    new XElement(PNs + "nvPr")),
                new XElement(
                    PNs + "grpSpPr",
                    new XElement(
                        ANs + "xfrm",
                        new XElement(
                            ANs + "off",
                            new XAttribute("x", offX),
                            new XAttribute("y", offY)),
                        new XElement(
                            ANs + "ext",
                            new XAttribute("cx", extCx),
                            new XAttribute("cy", extCy)),
                        new XElement(
                            ANs + "chOff",
                            new XAttribute("x", "0"),
                            new XAttribute("y", "0")),
                        new XElement(
                            ANs + "chExt",
                            new XAttribute("cx", extCx),
                            new XAttribute("cy", extCy)))));
        }

        private static int ReadOwnShapeId(XElement shapeEl)
        {
            if (shapeEl == null)
            {
                return 0;
            }

            foreach (XElement nv in shapeEl.Elements())
            {
                if (nv.Name.Namespace != PNs)
                {
                    continue;
                }

                string local = nv.Name.LocalName;
                if (local != "nvSpPr"
                    && local != "nvGrpSpPr"
                    && local != "nvPicPr"
                    && local != "nvCxnSpPr"
                    && local != "nvGraphicFramePr")
                {
                    continue;
                }

                XElement cNvPr = nv.Element(PNs + "cNvPr");
                if (cNvPr == null)
                {
                    continue;
                }

                string raw = (string)cNvPr.Attribute("id");
                int id;
                if (!string.IsNullOrEmpty(raw)
                    && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out id))
                {
                    return id;
                }
            }

            return 0;
        }

        private static int MaxCnvPrId(XElement root)
        {
            int max = 1;
            foreach (XElement cNvPr in root.Descendants(PNs + "cNvPr"))
            {
                string raw = (string)cNvPr.Attribute("id");
                int id;
                if (!string.IsNullOrEmpty(raw)
                    && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out id)
                    && id > max)
                {
                    max = id;
                }
            }

            return max;
        }

        private static string ToEmu(double points)
        {
            return ToEmuLong(points).ToString(CultureInfo.InvariantCulture);
        }

        private static long ToEmuLong(double points)
        {
            long emu = (long)Math.Round(points * EmuPerPoint);
            return emu < 0 ? 0 : emu;
        }

        private static void ShiftAllXfrmInTree(XElement root, long originX, long originY)
        {
            if (root == null || (originX == 0 && originY == 0))
            {
                return;
            }

            var xfrms = new List<XElement>();
            CollectXfrm(root, xfrms);
            for (int i = 0; i < xfrms.Count; i++)
            {
                XElement xfrm = xfrms[i];
                ShiftOffEl(xfrm.Element(ANs + "off") ?? xfrm.Element(PNs + "off"), originX, originY);
                ShiftOffEl(xfrm.Element(ANs + "chOff") ?? xfrm.Element(PNs + "chOff"), originX, originY);
            }
        }

        private static void CollectXfrm(XElement el, List<XElement> into)
        {
            if (el.Name.LocalName == "xfrm")
            {
                into.Add(el);
            }

            foreach (XElement child in el.Elements())
            {
                CollectXfrm(child, into);
            }
        }

        private static void ShiftOffEl(XElement off, long originX, long originY)
        {
            if (off == null)
            {
                return;
            }

            long x = ReadLongAttr(off, "x") - originX;
            long y = ReadLongAttr(off, "y") - originY;
            if (x < 0)
            {
                x = 0;
            }

            if (y < 0)
            {
                y = 0;
            }

            off.SetAttributeValue("x", x.ToString(CultureInfo.InvariantCulture));
            off.SetAttributeValue("y", y.ToString(CultureInfo.InvariantCulture));
        }

        private static void ShiftOwnXfrmOff(XElement shapeEl, long originX, long originY)
        {
            if (shapeEl == null)
            {
                return;
            }

            XElement xfrm = FindOwnXfrm(shapeEl);
            if (xfrm == null)
            {
                return;
            }

            ShiftOffEl(xfrm.Element(ANs + "off") ?? xfrm.Element(PNs + "off"), originX, originY);
        }

        private static XElement FindOwnXfrm(XElement shapeEl)
        {
            string local = shapeEl.Name.LocalName;
            if (local == "grpSp")
            {
                XElement grpSpPr = shapeEl.Element(PNs + "grpSpPr");
                return grpSpPr == null ? null : (grpSpPr.Element(ANs + "xfrm") ?? grpSpPr.Element(PNs + "xfrm"));
            }

            if (local == "graphicFrame")
            {
                return shapeEl.Element(PNs + "xfrm") ?? shapeEl.Element(ANs + "xfrm");
            }

            foreach (XElement child in shapeEl.Elements())
            {
                string name = child.Name.LocalName;
                if (name != "spPr" && name != "xfrm")
                {
                    continue;
                }

                if (name == "xfrm")
                {
                    return child;
                }

                XElement xfrm = child.Element(ANs + "xfrm") ?? child.Element(PNs + "xfrm");
                if (xfrm != null)
                {
                    return xfrm;
                }
            }

            return null;
        }

        private static long ReadLongAttr(XElement el, string name)
        {
            string raw = el == null ? null : (string)el.Attribute(name);
            long value;
            if (string.IsNullOrEmpty(raw)
                || !long.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return 0;
            }

            return value;
        }

        private static string DescribePptParents(PowerPoint.Shape group)
        {
            try
            {
                int outerId = group.Id;
                PowerPoint.GroupShapes items = group.GroupItems;
                var parts = new List<string>();
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape child = items[i];
                    string parentText = "-";
                    try
                    {
                        PowerPoint.Shape parent = child.ParentGroup;
                        parentText = parent == null
                            ? "null"
                            : (((int)parent.Type) + ":" + parent.Id + (parent.Id == outerId ? "=outer" : ""));
                    }
                    catch (Exception ex)
                    {
                        parentText = ex.Message;
                    }

                    parts.Add(child.Id + "->" + parentText);
                }

                return count + "[" + string.Join(",", parts.ToArray()) + "]";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private static bool TryReadPptBox(
            IList<PowerPoint.Shape> members,
            List<int> ids,
            out float left,
            out float top,
            out float width,
            out float height,
            out string error)
        {
            left = top = width = height = 0;
            error = null;
            float minL = float.MaxValue;
            float minT = float.MaxValue;
            float maxR = float.MinValue;
            float maxB = float.MinValue;
            for (int i = 0; i < members.Count; i++)
            {
                PowerPoint.Shape shape = members[i];
                try
                {
                    ids.Add(shape.Id);
                    float l = shape.Left;
                    float t = shape.Top;
                    float r = l + shape.Width;
                    float b = t + shape.Height;
                    if (l < minL)
                    {
                        minL = l;
                    }

                    if (t < minT)
                    {
                        minT = t;
                    }

                    if (r > maxR)
                    {
                        maxR = r;
                    }

                    if (b > maxB)
                    {
                        maxB = b;
                    }
                }
                catch (Exception ex)
                {
                    error = "OOXML 编组：读成员框失败: " + ex.Message;
                    return false;
                }
            }

            if (ids.Count < 2 || minL == float.MaxValue)
            {
                error = "OOXML 编组：成员框无效";
                return false;
            }

            left = minL;
            top = minT;
            width = Math.Max(1f, maxR - minL);
            height = Math.Max(1f, maxB - minT);
            return true;
        }

        private static bool TryReadWppBox(
            IList<object> members,
            List<int> ids,
            out double left,
            out double top,
            out double width,
            out double height,
            out string error)
        {
            left = top = width = height = 0;
            error = null;
            double minL = double.MaxValue;
            double minT = double.MaxValue;
            double maxR = double.MinValue;
            double maxB = double.MinValue;
            for (int i = 0; i < members.Count; i++)
            {
                object shape = members[i];
                try
                {
                    object rawId = WppCom.GetProperty(shape, "Id");
                    if (rawId == null)
                    {
                        error = "OOXML 编组：成员没有 Id";
                        return false;
                    }

                    ids.Add(Convert.ToInt32(rawId));
                    double l = Convert.ToDouble(WppCom.GetProperty(shape, "Left"));
                    double t = Convert.ToDouble(WppCom.GetProperty(shape, "Top"));
                    double r = l + Convert.ToDouble(WppCom.GetProperty(shape, "Width"));
                    double b = t + Convert.ToDouble(WppCom.GetProperty(shape, "Height"));
                    if (l < minL)
                    {
                        minL = l;
                    }

                    if (t < minT)
                    {
                        minT = t;
                    }

                    if (r > maxR)
                    {
                        maxR = r;
                    }

                    if (b > maxB)
                    {
                        maxB = b;
                    }
                }
                catch (Exception ex)
                {
                    error = "OOXML 编组：读成员框失败: " + ex.Message;
                    return false;
                }
            }

            if (ids.Count < 2 || minL == double.MaxValue)
            {
                error = "OOXML 编组：成员框无效";
                return false;
            }

            left = minL;
            top = minT;
            width = Math.Max(1.0, maxR - minL);
            height = Math.Max(1.0, maxB - minT);
            return true;
        }

        public static bool TryGetDirectGroupChildren(
            PowerPoint.Presentation presentation,
            int slideIndex,
            int groupId,
            out List<PptHtmlGroupXmlChild> children,
            out List<int> leafIds)
        {
            return TryGetDirectGroupChildrenCore(
                () => SaveCopyAsPpt(presentation),
                BuildPresKeyPpt(presentation),
                slideIndex,
                groupId,
                out children,
                out leafIds);
        }

        public static bool TryGetDirectGroupChildrenWpp(
            object presentation,
            int slideIndex,
            int groupId,
            out List<PptHtmlGroupXmlChild> children,
            out List<int> leafIds)
        {
            return TryGetDirectGroupChildrenCore(
                () =>
                {
                    string path = Path.Combine(
                        Path.GetTempPath(),
                        "ew-wpp-grptree-" + Guid.NewGuid().ToString("N") + ".pptx");
                    string saveError;
                    return TrySaveCopyAsWpp(presentation, path, out saveError) ? path : null;
                },
                BuildPresKeyWpp(presentation),
                slideIndex,
                groupId,
                out children,
                out leafIds);
        }

        private static bool TryCopyInnerGroupByUngroupPpt(
            PowerPoint.Slide slide,
            int groupId,
            out string error)
        {
            error = null;
            PowerPoint.Presentation pres;
            PowerPoint.Application app;
            try
            {
                pres = slide.Parent as PowerPoint.Presentation;
                app = pres == null ? null : pres.Application;
            }
            catch (Exception)
            {
                pres = null;
                app = null;
            }

            if (pres == null || app == null)
            {
                error = "页内找不到 ShapeId 组 " + groupId;
                return false;
            }

            int slideIndex;
            try
            {
                slideIndex = slide.SlideIndex;
            }
            catch (Exception)
            {
                error = "拷内组：读不到 SlideIndex";
                return false;
            }

            List<int> leafIds;
            if (!TryGetDirectGroupChildren(pres, slideIndex, groupId, out _, out leafIds)
                || leafIds == null
                || leafIds.Count == 0)
            {
                error = "页内找不到 ShapeId 组 " + groupId;
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "ew-ppt-dupinner-" + Guid.NewGuid().ToString("N") + ".pptx");
            PowerPoint.Presentation copy = null;
            try
            {
                pres.SaveCopyAs(tempPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                copy = TryOpenCopyPpt(app, tempPath, out error);
                if (copy == null)
                {
                    return false;
                }

                PowerPoint.Slide copySlide = copy.Slides[slideIndex];
                for (int step = 0; step < 8; step++)
                {
                    PowerPoint.Shape found = FindTopLevelPpt(copySlide.Shapes, groupId);
                    if (IsGroupShape(found))
                    {
                        found.Copy();
                        return true;
                    }

                    if (!UngroupAncestorPpt(copySlide.Shapes, groupId, leafIds))
                    {
                        break;
                    }
                }

                error = "无法把内组提升为顶层再拷贝 id=" + groupId;
                return false;
            }
            catch (Exception ex)
            {
                error = "拷内组失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        copy.Saved = Office.MsoTriState.msoTrue;
                        copy.Close();
                    }
                    catch (Exception)
                    {
                    }
                }

                TryDeleteFile(tempPath);
            }
        }

        private static bool UngroupAncestorPpt(
            PowerPoint.Shapes shapes,
            int groupId,
            List<int> leafIds)
        {
            if (shapes == null)
            {
                return false;
            }

            try
            {
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    if (!IsGroupShape(shape) || shape.Id == groupId)
                    {
                        continue;
                    }

                    if (!GroupOwnsLeavesPpt(shape, leafIds))
                    {
                        continue;
                    }

                    shape.Ungroup();
                    return true;
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool GroupOwnsLeavesPpt(PowerPoint.Shape group, List<int> leafIds)
        {
            try
            {
                PowerPoint.GroupShapes items = group.GroupItems;
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    int id = items[i].Id;
                    for (int j = 0; j < leafIds.Count; j++)
                    {
                        if (leafIds[j] == id)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool TryCopyInnerGroupByUngroupWpp(
            object slide,
            int groupId,
            out string error)
        {
            error = null;
            object pres = WppCom.GetProperty(slide, "Parent");
            object app = pres == null ? null : WppCom.GetProperty(pres, "Application");
            if (pres == null || app == null)
            {
                error = "页内找不到 ShapeId 组 " + groupId;
                return false;
            }

            int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
            List<int> leafIds;
            if (slideIndex < 1
                || !TryGetDirectGroupChildrenWpp(pres, slideIndex, groupId, out _, out leafIds)
                || leafIds == null
                || leafIds.Count == 0)
            {
                error = "页内找不到 ShapeId 组 " + groupId;
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "ew-wpp-dupinner-" + Guid.NewGuid().ToString("N") + ".pptx");
            object copy = null;
            try
            {
                if (!TrySaveCopyAsWpp(pres, tempPath, out error))
                {
                    return false;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                copy = TryOpenCopyWpp(presentations, tempPath, out error);
                if (copy == null)
                {
                    return false;
                }

                object copySlides = WppCom.GetProperty(copy, "Slides");
                object copySlide = WppCom.GetIndexed(copySlides, slideIndex);
                object copyShapes = copySlide == null ? null : WppCom.GetProperty(copySlide, "Shapes");
                for (int step = 0; step < 8; step++)
                {
                    object found = FindTopLevelWpp(copyShapes, groupId);
                    if (IsWppGroup(found))
                    {
                        WppCom.Invoke(found, "Copy");
                        return true;
                    }

                    if (!UngroupAncestorWpp(copyShapes, groupId, leafIds))
                    {
                        break;
                    }

                    copyShapes = WppCom.GetProperty(copySlide, "Shapes");
                }

                error = "无法把内组提升为顶层再拷贝 id=" + groupId;
                return false;
            }
            catch (Exception ex)
            {
                error = "拷内组失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (copy != null)
                {
                    try
                    {
                        WppCom.Invoke(copy, "Close");
                    }
                    catch (Exception)
                    {
                    }
                }

                TryDeleteFile(tempPath);
            }
        }

        private static bool UngroupAncestorWpp(object shapes, int groupId, List<int> leafIds)
        {
            if (shapes == null)
            {
                return false;
            }

            try
            {
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    if (!IsWppGroup(shape))
                    {
                        continue;
                    }

                    object rawId = WppCom.GetProperty(shape, "Id");
                    if (rawId != null && Convert.ToInt32(rawId) == groupId)
                    {
                        continue;
                    }

                    if (!GroupOwnsLeavesWpp(shape, leafIds))
                    {
                        continue;
                    }

                    WppCom.Invoke(shape, "Ungroup");
                    return true;
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool GroupOwnsLeavesWpp(object group, List<int> leafIds)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object child = WppCom.GetIndexed(items, i);
                    object rawId = child == null ? null : WppCom.GetProperty(child, "Id");
                    if (rawId == null)
                    {
                        continue;
                    }

                    int id = Convert.ToInt32(rawId);
                    for (int j = 0; j < leafIds.Count; j++)
                    {
                        if (leafIds[j] == id)
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        public static void InvalidateGroupTreeCache()
        {
            lock (TreeCacheLock)
            {
                TreeCacheKey = null;
                TreeCache = null;
                TreeCacheUtc = DateTime.MinValue;
            }
        }

        private static readonly object TreeCacheLock = new object();
        private static string TreeCacheKey;
        private static Dictionary<string, GroupXmlEntry> TreeCache;
        private static DateTime TreeCacheUtc;

        private sealed class GroupXmlEntry
        {
            public List<PptHtmlGroupXmlChild> Children = new List<PptHtmlGroupXmlChild>();
            public List<int> LeafIds = new List<int>();
        }

        private static bool TryGetDirectGroupChildrenCore(
            Func<string> saveCopyAs,
            string cacheKey,
            int slideIndex,
            int groupId,
            out List<PptHtmlGroupXmlChild> children,
            out List<int> leafIds)
        {
            children = null;
            leafIds = null;
            if (string.IsNullOrEmpty(cacheKey) || slideIndex < 1 || groupId < 1)
            {
                return false;
            }

            lock (TreeCacheLock)
            {
                if (TreeCache == null
                    || !string.Equals(TreeCacheKey, cacheKey, StringComparison.Ordinal)
                    || (DateTime.UtcNow - TreeCacheUtc).TotalSeconds >= 8)
                {
                    string tempPath = null;
                    try
                    {
                        tempPath = saveCopyAs();
                        if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
                        {
                            return false;
                        }

                        TreeCache = ParseGroupTree(tempPath);
                        TreeCacheKey = cacheKey;
                        TreeCacheUtc = DateTime.UtcNow;
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                    finally
                    {
                        TryDeleteFile(tempPath);
                    }
                }

                string mapKey = slideIndex.ToString(CultureInfo.InvariantCulture) + ":"
                    + groupId.ToString(CultureInfo.InvariantCulture);
                GroupXmlEntry entry;
                if (TreeCache == null || !TreeCache.TryGetValue(mapKey, out entry) || entry == null)
                {
                    return false;
                }

                children = entry.Children;
                leafIds = entry.LeafIds;
                return children != null && children.Count > 0;
            }
        }

        private static Dictionary<string, GroupXmlEntry> ParseGroupTree(string pptxPath)
        {
            var result = new Dictionary<string, GroupXmlEntry>(StringComparer.Ordinal);
            using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
            {
                List<string> slides = ListSlidePartPaths(zip);
                for (int s = 0; s < slides.Count; s++)
                {
                    ZipArchiveEntry entry = FindZipEntry(zip, slides[s]);
                    if (entry == null)
                    {
                        continue;
                    }

                    XDocument doc;
                    using (Stream stream = entry.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    if (doc.Root == null)
                    {
                        continue;
                    }

                    foreach (XElement grp in doc.Root.Descendants(PNs + "grpSp"))
                    {
                        IndexGrpSp(grp, s + 1, result);
                    }
                }
            }

            return result;
        }

        private static void IndexGrpSp(
            XElement grp,
            int slideIndex,
            Dictionary<string, GroupXmlEntry> into)
        {
            int id = ReadOwnShapeId(grp);
            if (id < 1)
            {
                return;
            }

            string key = slideIndex.ToString(CultureInfo.InvariantCulture) + ":"
                + id.ToString(CultureInfo.InvariantCulture);
            if (into.ContainsKey(key))
            {
                return;
            }

            var entry = new GroupXmlEntry();
            foreach (XElement child in grp.Elements())
            {
                string local = child.Name.LocalName;
                if (local == "nvGrpSpPr" || local == "grpSpPr" || local == "extLst")
                {
                    continue;
                }

                int childId = ReadOwnShapeId(child);
                if (childId < 1)
                {
                    continue;
                }

                bool isGroup = local == "grpSp";
                entry.Children.Add(new PptHtmlGroupXmlChild { Id = childId, IsGroup = isGroup });
                if (isGroup)
                {
                    IndexGrpSp(child, slideIndex, into);
                    string childKey = slideIndex.ToString(CultureInfo.InvariantCulture) + ":"
                        + childId.ToString(CultureInfo.InvariantCulture);
                    GroupXmlEntry nested;
                    if (into.TryGetValue(childKey, out nested) && nested != null)
                    {
                        entry.LeafIds.AddRange(nested.LeafIds);
                    }
                }
                else
                {
                    entry.LeafIds.Add(childId);
                }
            }

            into[key] = entry;
        }

        private static string SaveCopyAsPpt(PowerPoint.Presentation presentation)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-ppt-grptree-" + Guid.NewGuid().ToString("N") + ".pptx");
            presentation.SaveCopyAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            return path;
        }

        private static string BuildPresKeyPpt(PowerPoint.Presentation presentation)
        {
            try
            {
                return "ppt:" + (presentation.FullName ?? presentation.Name) + "@" + presentation.Slides.Count;
            }
            catch (Exception)
            {
                return "ppt:" + Guid.NewGuid().ToString("N");
            }
        }

        private static string BuildPresKeyWpp(object presentation)
        {
            try
            {
                string full = Convert.ToString(WppCom.GetProperty(presentation, "FullName")
                    ?? WppCom.GetProperty(presentation, "Name")
                    ?? "anon");
                object slides = WppCom.GetProperty(presentation, "Slides");
                int n = slides == null ? 0 : Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                return "wpp:" + full + "@" + n;
            }
            catch (Exception)
            {
                return "wpp:" + Guid.NewGuid().ToString("N");
            }
        }

        private static bool XmlGroupHasChildGroup(string pptxPath, int slideIndex, int groupId)
        {
            if (string.IsNullOrEmpty(pptxPath) || !File.Exists(pptxPath) || groupId < 1)
            {
                return false;
            }

            try
            {
                using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
                {
                    List<string> slides = ListSlidePartPaths(zip);
                    if (slideIndex < 1 || slideIndex > slides.Count)
                    {
                        return false;
                    }

                    ZipArchiveEntry entry = FindZipEntry(zip, slides[slideIndex - 1]);
                    if (entry == null)
                    {
                        return false;
                    }

                    XDocument doc;
                    using (Stream stream = entry.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    XElement grp = FindGrpSpById(doc.Root, groupId);
                    if (grp == null)
                    {
                        return false;
                    }

                    foreach (XElement child in grp.Elements())
                    {
                        if (child.Name == (PNs + "grpSp") || child.Name.LocalName == "grpSp")
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static XElement FindGrpSpById(XElement root, int groupId)
        {
            if (root == null)
            {
                return null;
            }

            foreach (XElement el in root.Descendants(PNs + "grpSp"))
            {
                if (ReadOwnShapeId(el) == groupId)
                {
                    return el;
                }
            }

            return null;
        }

        private static bool NestLooksAlivePpt(PowerPoint.Shape group)
        {
            if (group == null)
            {
                return false;
            }

            if (HasGroupItemPpt(group))
            {
                return true;
            }

            try
            {
                int outerId = group.Id;
                PowerPoint.GroupShapes items = group.GroupItems;
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        PowerPoint.Shape parent = items[i].ParentGroup;
                        if (parent != null && parent.Id != outerId && IsGroupShape(parent))
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool NestLooksAliveWpp(object group)
        {
            if (group == null)
            {
                return false;
            }

            if (HasGroupItemWpp(group))
            {
                return true;
            }

            try
            {
                int outerId = Convert.ToInt32(WppCom.GetProperty(group, "Id") ?? 0);
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        object child = WppCom.GetIndexed(items, i);
                        object parent = child == null ? null : WppCom.GetProperty(child, "ParentGroup");
                        object rawParentId = parent == null ? null : WppCom.GetProperty(parent, "Id");
                        if (rawParentId != null
                            && Convert.ToInt32(rawParentId) != outerId
                            && IsWppGroup(parent))
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool HasGroupItemPpt(PowerPoint.Shape group)
        {
            try
            {
                PowerPoint.GroupShapes items = group.GroupItems;
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    if (IsGroupShape(items[i]))
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static bool HasGroupItemWpp(object group)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    if (IsWppGroup(WppCom.GetIndexed(items, i)))
                    {
                        return true;
                    }
                }
            }
            catch (Exception)
            {
            }

            return false;
        }

        private static string DescribePptGroupItems(PowerPoint.Shape group)
        {
            try
            {
                PowerPoint.GroupShapes items = group.GroupItems;
                var parts = new List<string>();
                int count = items.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape child = items[i];
                    parts.Add(((int)child.Type) + ":" + child.Id);
                }

                return count + "[" + string.Join(",", parts.ToArray()) + "]";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private static string DescribeWppGroupItems(object group)
        {
            try
            {
                object items = WppCom.GetProperty(group, "GroupItems");
                int count = Convert.ToInt32(WppCom.GetProperty(items, "Count") ?? 0);
                var parts = new List<string>();
                for (int i = 1; i <= count; i++)
                {
                    object child = WppCom.GetIndexed(items, i);
                    object type = child == null ? null : WppCom.GetProperty(child, "Type");
                    object rawId = child == null ? null : WppCom.GetProperty(child, "Id");
                    parts.Add((type ?? "?") + ":" + (rawId ?? "-"));
                }

                return count + "[" + string.Join(",", parts.ToArray()) + "]";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        private static PowerPoint.Shape FindTopLevelPpt(PowerPoint.Shapes shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                int count = shapes.Count;
                for (int i = 1; i <= count; i++)
                {
                    PowerPoint.Shape shape = shapes[i];
                    if (shape.Id == id)
                    {
                        return shape;
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static object FindTopLevelWpp(object shapes, int id)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                int count = Convert.ToInt32(WppCom.GetProperty(shapes, "Count") ?? 0);
                for (int i = 1; i <= count; i++)
                {
                    object shape = WppCom.GetIndexed(shapes, i);
                    object rawId = shape == null ? null : WppCom.GetProperty(shape, "Id");
                    if (rawId != null && Convert.ToInt32(rawId) == id)
                    {
                        return shape;
                    }
                }
            }
            catch (Exception)
            {
            }

            return null;
        }

        private static PowerPoint.Presentation TryOpenCopyPpt(
            PowerPoint.Application app,
            string path,
            out string error)
        {
            error = null;
            try
            {
                return app.Presentations.Open(
                    path,
                    Office.MsoTriState.msoTrue,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                error = "OOXML 编组：打开副本失败: " + ex.Message;
                return null;
            }
        }

        private static object TryOpenCopyWpp(object presentations, string path, out string error)
        {
            error = null;
            if (presentations == null)
            {
                error = "OOXML 编组：没有 Presentations";
                return null;
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", path);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", path, false, false, false);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", path, true);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception ex)
            {
                error = "OOXML 编组：打开副本失败: " + ex.Message;
            }

            if (error == null)
            {
                error = "OOXML 编组：打开副本无返回";
            }

            return null;
        }

        private static object TryPasteWpp(object shapes)
        {
            if (shapes == null)
            {
                return null;
            }

            try
            {
                object pasted = WppCom.Invoke(shapes, "Paste");
                if (pasted != null)
                {
                    object first = WppCom.GetIndexed(pasted, 1);
                    return first ?? pasted;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object pasted = WppCom.Invoke(shapes, "PasteSpecial", 0);
                object first = pasted == null ? null : WppCom.GetIndexed(pasted, 1);
                return first ?? pasted;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TrySaveCopyAsWpp(object pres, string path, out string error)
        {
            error = null;
            try
            {
                WppCom.Invoke(pres, "SaveCopyAs", path);
            }
            catch (Exception)
            {
                try
                {
                    WppCom.Invoke(pres, "SaveCopyAs", path, 24);
                }
                catch (Exception ex)
                {
                    error = "OOXML 编组：SaveCopyAs 失败: " + ex.Message;
                    return false;
                }
            }

            if (!File.Exists(path))
            {
                error = "OOXML 编组：SaveCopyAs 未落盘";
                return false;
            }

            return true;
        }

        private static bool TrySilenceWppAlerts(object app, out object prev)
        {
            prev = null;
            try
            {
                prev = WppCom.GetProperty(app, "DisplayAlerts");
                WppCom.TrySetProperty(app, "DisplayAlerts", false);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void TryRestoreWppAlerts(object app, object prev)
        {
            if (app == null)
            {
                return;
            }

            try
            {
                WppCom.TrySetProperty(app, "DisplayAlerts", prev);
            }
            catch (Exception)
            {
            }
        }

        private static void TryDeletePpt(PowerPoint.Shape shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                shape.Delete();
            }
            catch (Exception)
            {
            }
        }

        private static void TryDeleteWpp(object shape)
        {
            if (shape == null)
            {
                return;
            }

            try
            {
                WppCom.Invoke(shape, "Delete");
            }
            catch (Exception)
            {
            }
        }

        private static void TryDeleteFile(string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        private static List<string> ListSlidePartPaths(ZipArchive zip)
        {
            ZipArchiveEntry pres = FindZipEntry(zip, "ppt/presentation.xml");
            ZipArchiveEntry rels = FindZipEntry(zip, "ppt/_rels/presentation.xml.rels");
            if (pres == null || rels == null)
            {
                var fallback = new List<string>();
                foreach (ZipArchiveEntry e in zip.Entries)
                {
                    string p = e.FullName.Replace('\\', '/');
                    if (p.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                        && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        && p.IndexOf("_rels", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        fallback.Add(p);
                    }
                }

                fallback.Sort(StringComparer.OrdinalIgnoreCase);
                return fallback;
            }

            XDocument relDoc;
            using (Stream s = rels.Open())
            {
                relDoc = XDocument.Load(s);
            }

            var ridToTarget = new Dictionary<string, string>(StringComparer.Ordinal);
            if (relDoc.Root != null)
            {
                foreach (XElement rel in relDoc.Root.Elements(RelPkgNs + "Relationship"))
                {
                    string id = (string)rel.Attribute("Id");
                    string target = (string)rel.Attribute("Target");
                    if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target))
                    {
                        continue;
                    }

                    string t = target.Replace('\\', '/');
                    if (!t.StartsWith("/", StringComparison.Ordinal)
                        && !t.StartsWith("ppt/", StringComparison.OrdinalIgnoreCase))
                    {
                        t = "ppt/" + t.TrimStart('/');
                    }
                    else if (t.StartsWith("/", StringComparison.Ordinal))
                    {
                        t = t.TrimStart('/');
                    }

                    ridToTarget[id] = t;
                }
            }

            XDocument presDoc;
            using (Stream s = pres.Open())
            {
                presDoc = XDocument.Load(s);
            }

            var ordered = new List<string>();
            XElement sldIdLst = presDoc.Root == null ? null : presDoc.Root.Element(PNs + "sldIdLst");
            if (sldIdLst == null)
            {
                foreach (string v in ridToTarget.Values)
                {
                    if (v.IndexOf("/slides/slide", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        ordered.Add(v);
                    }
                }

                ordered.Sort(StringComparer.OrdinalIgnoreCase);
                return ordered;
            }

            foreach (XElement sldId in sldIdLst.Elements(PNs + "sldId"))
            {
                string rid = (string)sldId.Attribute(RNs + "id");
                string path;
                if (!string.IsNullOrEmpty(rid) && ridToTarget.TryGetValue(rid, out path))
                {
                    ordered.Add(path);
                }
            }

            return ordered;
        }

        private static ZipArchiveEntry FindZipEntry(ZipArchive zip, string fullName)
        {
            ZipArchiveEntry e = zip.GetEntry(fullName);
            if (e != null)
            {
                return e;
            }

            string want = fullName.Replace('\\', '/');
            foreach (ZipArchiveEntry x in zip.Entries)
            {
                if (string.Equals(x.FullName.Replace('\\', '/'), want, StringComparison.OrdinalIgnoreCase))
                {
                    return x;
                }
            }

            return null;
        }
    }
}
