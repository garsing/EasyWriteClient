using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    /// <summary>从 pptx 包读/改幻灯片背景图（blipFill）。不打开演示文稿。</summary>
    internal static class PptBackgroundPackage
    {
        internal const string ExportFolderName = "ppt_backgrounds";

        private static readonly XNamespace ANs =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

        internal sealed class PicturePlacement
        {
            public string Fit;
            public double StretchLeft;
            public double StretchRight;
            public double StretchTop;
            public double StretchBottom;
            public double TileOffsetX;
            public double TileOffsetY;
            public double TileScaleX = 100;
            public double TileScaleY = 100;
            public string TileAlign = "tl";
            public string TileMirror = "none";
            public string RelativePath;
        }

        public static bool TryReadPicture(
            string pptxPath,
            int slideIndex,
            string exportDirectory,
            out PicturePlacement placement,
            out string error)
        {
            placement = null;
            error = null;
            if (string.IsNullOrEmpty(pptxPath) || !File.Exists(pptxPath))
            {
                error = "背景图：包不存在";
                return false;
            }

            if (slideIndex < 1)
            {
                error = "背景图：页序无效";
                return false;
            }

            try
            {
                using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
                {
                    if (!TryFindSlidePart(zip, slideIndex, out string slidePart, out error))
                    {
                        return false;
                    }

                    XDocument slideDoc = LoadXml(zip, slidePart);
                    if (!TryFindBlipFill(slideDoc, out XElement blipFill, out error))
                    {
                        return false;
                    }

                    placement = ReadPlacement(blipFill);
                    if (!string.IsNullOrEmpty(exportDirectory))
                    {
                        if (!TryExportBlip(
                                zip,
                                slidePart,
                                blipFill,
                                exportDirectory,
                                out string relative,
                                out error))
                        {
                            placement = null;
                            return false;
                        }

                        placement.RelativePath = relative;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "背景图读取失败: " + ex.Message;
                placement = null;
                return false;
            }
        }

        public static bool TryRewritePlacement(
            string pptxPath,
            int slideIndex,
            PicturePlacement placement,
            out string error)
        {
            error = null;
            if (placement == null)
            {
                error = "背景图：摆放为空";
                return false;
            }

            string rewriteError = null;
            bool ok = PptHtmlOoxmlIo.TryRewriteSavedFile(
                pptxPath,
                (zip, replacements) =>
                {
                    if (!TryFindSlidePart(zip, slideIndex, out string slidePart, out rewriteError))
                    {
                        return false;
                    }

                    XDocument doc = LoadXml(zip, slidePart);
                    if (!TryFindBlipFill(doc, out XElement blipFill, out rewriteError))
                    {
                        return false;
                    }

                    WritePlacement(blipFill, placement);
                    replacements[slidePart] = ToUtf8(doc);
                    return true;
                },
                out error,
                "背景摆放");
            if (!ok && !string.IsNullOrEmpty(rewriteError))
            {
                error = rewriteError;
            }

            return ok;
        }

        public static bool Matches(PicturePlacement actual, PicturePlacement expected, out string detail)
        {
            detail = null;
            if (actual == null || expected == null)
            {
                detail = "摆放为空";
                return false;
            }

            if (!string.Equals(actual.Fit, expected.Fit, StringComparison.Ordinal))
            {
                detail = "fit 读回 " + actual.Fit + "，写入 " + expected.Fit;
                return false;
            }

            if (string.Equals(expected.Fit, "tile", StringComparison.Ordinal))
            {
                if (!Near(actual.TileOffsetX, expected.TileOffsetX, 0.5)
                    || !Near(actual.TileOffsetY, expected.TileOffsetY, 0.5))
                {
                    detail = "平铺偏移读回不一致";
                    return false;
                }

                if (!Near(actual.TileScaleX, expected.TileScaleX, 0.5)
                    || !Near(actual.TileScaleY, expected.TileScaleY, 0.5))
                {
                    detail = "平铺缩放读回 "
                    + actual.TileScaleX.ToString("0.###", CultureInfo.InvariantCulture)
                    + ","
                    + actual.TileScaleY.ToString("0.###", CultureInfo.InvariantCulture)
                    + " 期望 "
                    + expected.TileScaleX.ToString("0.###", CultureInfo.InvariantCulture)
                    + ","
                    + expected.TileScaleY.ToString("0.###", CultureInfo.InvariantCulture);
                    return false;
                }

                if (!string.Equals(actual.TileAlign, expected.TileAlign, StringComparison.Ordinal)
                    || !string.Equals(actual.TileMirror, expected.TileMirror, StringComparison.Ordinal))
                {
                    detail = "平铺对齐或镜像读回 align="
                    + (actual.TileAlign ?? "")
                    + " mirror="
                    + (actual.TileMirror ?? "")
                    + " 期望 align="
                    + (expected.TileAlign ?? "")
                    + " mirror="
                    + (expected.TileMirror ?? "");
                    return false;
                }

                return true;
            }

            if (!Near(actual.StretchLeft, expected.StretchLeft, 0.1)
                || !Near(actual.StretchRight, expected.StretchRight, 0.1)
                || !Near(actual.StretchTop, expected.StretchTop, 0.1)
                || !Near(actual.StretchBottom, expected.StretchBottom, 0.1))
            {
                detail = "拉伸偏移读回不一致";
                return false;
            }

            return true;
        }

        private static bool Near(double actual, double expected, double tol)
        {
            return Math.Abs(actual - expected) <= tol;
        }

        private static PicturePlacement ReadPlacement(XElement blipFill)
        {
            XElement tile = Child(blipFill, "tile");
            if (tile != null)
            {
                return new PicturePlacement
                {
                    Fit = "tile",
                    TileOffsetX = EmuToPoint(AttrLong(tile, "tx")),
                    TileOffsetY = EmuToPoint(AttrLong(tile, "ty")),
                    TileScaleX = Pct1000ToPercent(AttrLong(tile, "sx", 100000)),
                    TileScaleY = Pct1000ToPercent(AttrLong(tile, "sy", 100000)),
                    TileAlign = AttrText(tile, "algn", "tl"),
                    TileMirror = AttrText(tile, "flip", "none")
                };
            }

            XElement fillRect = Child(Child(blipFill, "stretch"), "fillRect");
            return new PicturePlacement
            {
                Fit = "stretch",
                StretchLeft = Pct1000ToPercent(AttrLong(fillRect, "l")),
                StretchRight = Pct1000ToPercent(AttrLong(fillRect, "r")),
                StretchTop = Pct1000ToPercent(AttrLong(fillRect, "t")),
                StretchBottom = Pct1000ToPercent(AttrLong(fillRect, "b"))
            };
        }

        private static void WritePlacement(XElement blipFill, PicturePlacement placement)
        {
            XElement stretch = Child(blipFill, "stretch");
            XElement tile = Child(blipFill, "tile");
            if (string.Equals(placement.Fit, "tile", StringComparison.Ordinal))
            {
                stretch?.Remove();
                if (tile == null)
                {
                    tile = new XElement(ANs + "tile");
                    blipFill.Add(tile);
                }

                SetAttr(tile, "tx", PointToEmu(placement.TileOffsetX).ToString(CultureInfo.InvariantCulture));
                SetAttr(tile, "ty", PointToEmu(placement.TileOffsetY).ToString(CultureInfo.InvariantCulture));
                SetAttr(tile, "sx", PercentToPct1000(placement.TileScaleX).ToString(CultureInfo.InvariantCulture));
                SetAttr(tile, "sy", PercentToPct1000(placement.TileScaleY).ToString(CultureInfo.InvariantCulture));
                SetAttr(tile, "flip", string.IsNullOrEmpty(placement.TileMirror) ? "none" : placement.TileMirror);
                SetAttr(tile, "algn", string.IsNullOrEmpty(placement.TileAlign) ? "tl" : placement.TileAlign);
                return;
            }

            tile?.Remove();
            if (stretch == null)
            {
                stretch = new XElement(ANs + "stretch");
                blipFill.Add(stretch);
            }

            XElement fillRect = Child(stretch, "fillRect");
            if (fillRect == null)
            {
                fillRect = new XElement(ANs + "fillRect");
                stretch.Add(fillRect);
            }

            SetAttr(fillRect, "l", PercentToPct1000(placement.StretchLeft).ToString(CultureInfo.InvariantCulture));
            SetAttr(fillRect, "t", PercentToPct1000(placement.StretchTop).ToString(CultureInfo.InvariantCulture));
            SetAttr(fillRect, "r", PercentToPct1000(placement.StretchRight).ToString(CultureInfo.InvariantCulture));
            SetAttr(fillRect, "b", PercentToPct1000(placement.StretchBottom).ToString(CultureInfo.InvariantCulture));
        }

        private static bool TryExportBlip(
            ZipArchive zip,
            string slidePart,
            XElement blipFill,
            string exportDirectory,
            out string relative,
            out string error)
        {
            relative = null;
            error = null;
            XElement blip = Child(blipFill, "blip");
            string embed = blip == null ? null : AttrText(blip, "embed", null);
            if (string.IsNullOrEmpty(embed))
            {
                error = "背景图：没有可导出的图片";
                return false;
            }

            string relsPart = RelsPartFor(slidePart);
            XDocument rels = LoadXml(zip, relsPart);
            XElement rel = rels.Root?.Elements().FirstOrDefault(e =>
                string.Equals(AttrText(e, "Id", null), embed, StringComparison.Ordinal));
            string target = rel == null ? null : AttrText(rel, "Target", null);
            if (string.IsNullOrEmpty(target))
            {
                error = "背景图：找不到图片关系 " + embed;
                return false;
            }

            string mediaPart = CombineZip(ParentPart(slidePart), target);
            ZipArchiveEntry entry = FindEntry(zip, mediaPart);
            if (entry == null)
            {
                error = "背景图：包内没有 " + mediaPart;
                return false;
            }

            byte[] bytes;
            using (Stream stream = entry.Open())
            using (var ms = new MemoryStream())
            {
                stream.CopyTo(ms);
                bytes = ms.ToArray();
            }

            if (bytes.Length == 0)
            {
                error = "背景图：图片为空";
                return false;
            }

            string ext = SafeExtension(target);
            string hash8;
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                var sb = new StringBuilder(8);
                for (int i = 0; i < hash.Length && sb.Length < 8; i++)
                {
                    sb.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                hash8 = sb.ToString(0, 8);
            }

            Directory.CreateDirectory(exportDirectory);
            string fileName = hash8 + ext;
            string full = Path.Combine(exportDirectory, fileName);
            if (!File.Exists(full))
            {
                File.WriteAllBytes(full, bytes);
            }

            relative = ExportFolderName + "/" + fileName;
            return true;
        }

        private static bool TryFindSlidePart(
            ZipArchive zip,
            int slideIndex,
            out string slidePart,
            out string error)
        {
            slidePart = null;
            error = null;
            XDocument pres = LoadXml(zip, "ppt/presentation.xml");
            XElement list = pres.Descendants().FirstOrDefault(e => e.Name.LocalName == "sldIdLst");
            if (list == null)
            {
                error = "背景图：包内没有幻灯片列表";
                return false;
            }

            List<XElement> ids = list.Elements().Where(e => e.Name.LocalName == "sldId").ToList();
            if (slideIndex > ids.Count)
            {
                error = "背景图：页序超出包内页数";
                return false;
            }

            string rid = AttrText(ids[slideIndex - 1], "id", null);
            // r:id 才是关系号。本地名 id 可能是 sldId 自己的 id。
            XAttribute ridAttr = ids[slideIndex - 1].Attributes()
                .FirstOrDefault(a => a.Name.LocalName == "id" && a.Name.NamespaceName.IndexOf("relationships", StringComparison.Ordinal) >= 0);
            if (ridAttr != null)
            {
                rid = ridAttr.Value;
            }

            if (string.IsNullOrEmpty(rid))
            {
                error = "背景图：幻灯片没有关系 Id";
                return false;
            }

            XDocument rels = LoadXml(zip, "ppt/_rels/presentation.xml.rels");
            XElement rel = rels.Root?.Elements().FirstOrDefault(e =>
                string.Equals(AttrText(e, "Id", null), rid, StringComparison.Ordinal));
            string target = rel == null ? null : AttrText(rel, "Target", null);
            if (string.IsNullOrEmpty(target))
            {
                error = "背景图：找不到幻灯片部件";
                return false;
            }

            slidePart = CombineZip("ppt", target);
            if (FindEntry(zip, slidePart) == null)
            {
                error = "背景图：缺少 " + slidePart;
                return false;
            }

            return true;
        }

        private static bool TryFindBlipFill(XDocument slideDoc, out XElement blipFill, out string error)
        {
            blipFill = null;
            error = null;
            XElement bg = slideDoc.Descendants().FirstOrDefault(e => e.Name.LocalName == "bg");
            XElement bgPr = bg == null ? null : Child(bg, "bgPr");
            blipFill = bgPr == null ? null : Child(bgPr, "blipFill");
            if (blipFill == null)
            {
                error = "背景图：这一页背景里没有图片";
                return false;
            }

            return true;
        }

        private static XElement Child(XElement parent, string local)
        {
            if (parent == null || string.IsNullOrEmpty(local))
            {
                return null;
            }

            return parent.Elements().FirstOrDefault(e => e.Name.LocalName == local);
        }

        private static string AttrText(XElement el, string local, string fallback)
        {
            if (el == null)
            {
                return fallback;
            }

            XAttribute attr = el.Attributes().FirstOrDefault(a => a.Name.LocalName == local);
            if (attr == null || string.IsNullOrEmpty(attr.Value))
            {
                return fallback;
            }

            return attr.Value.Trim();
        }

        private static long AttrLong(XElement el, string local, long fallback = 0)
        {
            string text = AttrText(el, local, null);
            if (string.IsNullOrEmpty(text))
            {
                return fallback;
            }

            if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long value))
            {
                return value;
            }

            return fallback;
        }

        private static void SetAttr(XElement el, string local, string value)
        {
            XAttribute attr = el.Attributes().FirstOrDefault(a => a.Name.LocalName == local);
            if (attr != null)
            {
                attr.Value = value;
                return;
            }

            el.SetAttributeValue(local, value);
        }

        private static double Pct1000ToPercent(long raw)
        {
            return raw / 1000.0;
        }

        private static long PercentToPct1000(double percent)
        {
            return (long)Math.Round(percent * 1000.0, MidpointRounding.AwayFromZero);
        }

        private static double EmuToPoint(long emu)
        {
            return emu / 12700.0;
        }

        private static long PointToEmu(double point)
        {
            return (long)Math.Round(point * 12700.0, MidpointRounding.AwayFromZero);
        }

        private static string SafeExtension(string target)
        {
            string ext = Path.GetExtension(target ?? "") ?? "";
            ext = ext.ToLowerInvariant();
            if (ext.Length >= 2 && ext.Length <= 9)
            {
                bool ok = true;
                for (int i = 1; i < ext.Length; i++)
                {
                    char c = ext[i];
                    if (!((c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')))
                    {
                        ok = false;
                        break;
                    }
                }

                if (ok)
                {
                    return ext;
                }
            }

            return ".png";
        }

        private static string RelsPartFor(string part)
        {
            string parent = ParentPart(part);
            string file = part.Substring(parent.Length).TrimStart('/');
            return parent + "/_rels/" + file + ".rels";
        }

        private static string ParentPart(string part)
        {
            int slash = part.LastIndexOf('/');
            if (slash < 0)
            {
                return "";
            }

            return part.Substring(0, slash);
        }

        private static string CombineZip(string baseDir, string relative)
        {
            var stack = new List<string>();
            if (!string.IsNullOrEmpty(baseDir))
            {
                foreach (string part in baseDir.Replace('\\', '/').Split('/'))
                {
                    if (!string.IsNullOrEmpty(part))
                    {
                        stack.Add(part);
                    }
                }
            }

            foreach (string part in (relative ?? "").Replace('\\', '/').Split('/'))
            {
                if (part.Length == 0 || part == ".")
                {
                    continue;
                }

                if (part == "..")
                {
                    if (stack.Count > 0)
                    {
                        stack.RemoveAt(stack.Count - 1);
                    }

                    continue;
                }

                stack.Add(part);
            }

            return string.Join("/", stack);
        }

        private static XDocument LoadXml(ZipArchive zip, string part)
        {
            ZipArchiveEntry entry = FindEntry(zip, part);
            if (entry == null)
            {
                throw new InvalidOperationException("包内没有 " + part);
            }

            using (Stream stream = entry.Open())
            {
                return XDocument.Load(stream);
            }
        }

        private static ZipArchiveEntry FindEntry(ZipArchive zip, string part)
        {
            string want = (part ?? "").Replace('\\', '/').TrimStart('/');
            foreach (ZipArchiveEntry entry in zip.Entries)
            {
                string name = (entry.FullName ?? "").Replace('\\', '/').TrimStart('/');
                if (string.Equals(name, want, StringComparison.OrdinalIgnoreCase))
                {
                    return entry;
                }
            }

            return null;
        }

        private static byte[] ToUtf8(XDocument doc)
        {
            var settings = new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false),
                OmitXmlDeclaration = false,
                CloseOutput = false
            };
            using (var ms = new MemoryStream())
            {
                using (XmlWriter writer = XmlWriter.Create(ms, settings))
                {
                    doc.Save(writer);
                }

                return ms.ToArray();
            }
        }
    }
}
