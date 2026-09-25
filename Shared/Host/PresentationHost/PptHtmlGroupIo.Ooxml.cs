using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Xml.Linq;
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

                int newGroupId = 0;
                string wrapErr = null;
                string pasteErr = null;
                PowerPoint.Shape createdGroup = null;
                if (!PptHtmlOoxmlIo.TryRewriteAndCopyBack(
                    pres,
                    slideIndex,
                    (zip, parts) => TryWrapMembersInPackage(
                        zip,
                        parts,
                        slideIndex,
                        ids,
                        left,
                        top,
                        width,
                        height,
                        out newGroupId,
                        out wrapErr),
                    (copyPres, openSlide) =>
                    {
                        PowerPoint.Presentation copy = copyPres as PowerPoint.Presentation;
                        if (copy == null)
                        {
                            return PptHtmlOoxmlIo.TryFindShapeById(copyPres, openSlide, newGroupId);
                        }

                        return FindTopLevelPpt(copy.Slides[openSlide].Shapes, newGroupId);
                    },
                    () =>
                    {
                        PowerPoint.ShapeRange pasted = slide.Shapes.Paste();
                        if (pasted == null || pasted.Count < 1)
                        {
                            pasteErr = "OOXML 编组：贴回失败";
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

                        for (int i = 0; i < members.Count; i++)
                        {
                            TryDeletePpt(members[i]);
                        }

                        InvalidateGroupTreeCache();
                        createdGroup = created;
                        return true;
                    },
                    out string pipeErr,
                    tag: "编组"))
                {
                    error = wrapErr ?? pasteErr ?? pipeErr;
                    return false;
                }

                group = createdGroup;
                return true;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组失败: " + ex.Message;
                return false;
            }
            finally
            {
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

            object prevAlerts = null;
            bool alertsSet = false;
            try
            {
                alertsSet = TrySilenceWppAlerts(app, out prevAlerts);
                int newGroupId = 0;
                string wrapErr = null;
                string pasteErr = null;
                object createdGroup = null;
                if (!PptHtmlOoxmlIo.TryRewriteAndCopyBack(
                    pres,
                    slideIndex,
                    (zip, parts) => TryWrapMembersInPackage(
                        zip,
                        parts,
                        slideIndex,
                        ids,
                        left,
                        top,
                        width,
                        height,
                        out newGroupId,
                        out wrapErr),
                    (copyPres, openSlide) =>
                    {
                        object copySlides = WppCom.GetProperty(copyPres, "Slides");
                        object copySlide = WppCom.GetIndexed(copySlides, openSlide);
                        object copyShapes = copySlide == null ? null : WppCom.GetProperty(copySlide, "Shapes");
                        return FindTopLevelWpp(copyShapes, newGroupId);
                    },
                    () =>
                    {
                        object shapes = WppCom.GetProperty(slide, "Shapes");
                        object pasted = TryPasteWpp(shapes);
                        if (pasted == null)
                        {
                            pasteErr = "OOXML 编组：贴回失败";
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

                        for (int i = 0; i < members.Count; i++)
                        {
                            TryDeleteWpp(members[i]);
                        }

                        InvalidateGroupTreeCache();
                        createdGroup = pasted;
                        return true;
                    },
                    out string pipeErr,
                    tag: "编组"))
                {
                    error = wrapErr ?? pasteErr ?? pipeErr;
                    return false;
                }

                group = createdGroup;
                return true;
            }
            catch (Exception ex)
            {
                error = "OOXML 编组失败: " + ex.Message;
                return false;
            }
            finally
            {
                if (alertsSet)
                {
                    TryRestoreWppAlerts(app, prevAlerts);
                }
            }
        }

        private static bool TryWrapMembersInPackage(
            ZipArchive zip,
            IDictionary<string, byte[]> replacements,
            int slideIndex,
            IList<int> memberIds,
            double left,
            double top,
            double width,
            double height,
            out int newGroupId,
            out string error)
        {
            newGroupId = 0;
            error = null;
            try
            {
                if (zip == null || replacements == null)
                {
                    error = "OOXML 编组：zip 空";
                    return false;
                }

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
                    out error))
                {
                    return false;
                }

                PptHtmlOoxmlIo.PutReplacement(replacements, entry.FullName, PptHtmlOoxmlIo.XmlPartBytes(doc));
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
            out string error)
        {
            newGroupId = 0;
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
                () =>
                {
                    string saveError;
                    return PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, "grptree", out saveError);
                },
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
                    string saveError;
                    return PptHtmlOoxmlIo.TrySaveCopyTemp(presentation, "grptree", out saveError);
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

            string tempPath = PptHtmlOoxmlIo.NewTempPath("dupinner");
            PowerPoint.Presentation copy = null;
            try
            {
                if (!PptHtmlOoxmlIo.TrySaveCopy(pres, tempPath, out error))
                {
                    return false;
                }

                copy = PptHtmlOoxmlIo.TryOpenCopy(app, tempPath, out error) as PowerPoint.Presentation;
                if (copy == null)
                {
                    if (string.IsNullOrEmpty(error))
                    {
                        error = "拷内组：打开副本失败";
                    }

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
                PptHtmlOoxmlIo.TryCloseCopy(copy);
                PptHtmlOoxmlIo.TryDeleteFile(tempPath);
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

            string tempPath = PptHtmlOoxmlIo.NewTempPath("dupinner");
            object copy = null;
            try
            {
                if (!PptHtmlOoxmlIo.TrySaveCopy(pres, tempPath, out error))
                {
                    return false;
                }

                copy = PptHtmlOoxmlIo.TryOpenCopy(app, tempPath, out error);
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
                PptHtmlOoxmlIo.TryCloseCopy(copy);
                PptHtmlOoxmlIo.TryDeleteFile(tempPath);
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
                        PptHtmlOoxmlIo.TryDeleteFile(tempPath);
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
