using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    internal static partial class PptHtmlOoxmlIo
    {
        private static readonly XNamespace RelNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace PNs =
            "http://schemas.openxmlformats.org/presentationml/2006/main";
        private static readonly XNamespace RNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace CtNs =
            "http://schemas.openxmlformats.org/package/2006/content-types";

        internal static void PutReplacement(IDictionary<string, byte[]> replacements, string part, byte[] data)
        {
            if (replacements == null || string.IsNullOrEmpty(part) || data == null)
            {
                return;
            }

            replacements[NormPart(part)] = data;
        }

        internal static byte[] XmlPartBytes(XDocument doc)
        {
            using (var ms = new MemoryStream())
            {
                doc.Save(ms);
                return ms.ToArray();
            }
        }

        private static bool TrySlimToSlideAndTheme(
            string srcPath,
            int slideNo,
            ZipArchive zin,
            IDictionary<string, byte[]> replacements,
            out string slimPath,
            out string error)
        {
            slimPath = null;
            error = null;
            if (string.IsNullOrEmpty(srcPath) || zin == null || slideNo < 1)
            {
                error = "瘦包参数无效";
                return false;
            }

            slimPath = srcPath + ".slim.pptx";
            TryDeleteFile(slimPath);
            try
            {
                var map = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry e in zin.Entries)
                {
                    if (string.IsNullOrEmpty(e.Name) && e.FullName.EndsWith("/"))
                    {
                        continue;
                    }

                    map[NormPart(e.FullName)] = e;
                }

                string slidePart;
                if (!TryResolveSlidePart(map, replacements, slideNo, out slidePart, out error))
                {
                    return false;
                }

                HashSet<string> need = CollectSlideThemeClosure(map, replacements, slidePart);
                if (need.Count < 4)
                {
                    error = "瘦包闭包太小";
                    return false;
                }

                byte[] presXml = RewritePresentationOneSlide(
                    ReadPart(map, replacements, "ppt/presentation.xml"), slidePart, map, replacements);
                byte[] presRels = RewritePresentationRels(
                    ReadPart(map, replacements, "ppt/_rels/presentation.xml.rels"), need);
                byte[] types = RewriteContentTypes(
                    ReadPart(map, replacements, "[Content_Types].xml"), need);
                if (presXml == null || presRels == null || types == null)
                {
                    error = "瘦包改 presentation/Content_Types 失败";
                    return false;
                }

                using (FileStream fs = File.Create(slimPath))
                using (var zout = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    foreach (string part in need.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
                    {
                        if (part.Equals("ppt/presentation.xml", StringComparison.OrdinalIgnoreCase)
                            || part.Equals("ppt/_rels/presentation.xml.rels", StringComparison.OrdinalIgnoreCase)
                            || part.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        byte[] overlay = LookupReplacement(replacements, part);
                        if (overlay != null)
                        {
                            WriteZipBytes(zout, part, overlay);
                            continue;
                        }

                        ZipArchiveEntry src;
                        if (!map.TryGetValue(part, out src))
                        {
                            continue;
                        }

                        CopyZipEntry(src, zout, part);
                    }

                    WriteZipBytes(zout, "ppt/presentation.xml", presXml);
                    WriteZipBytes(zout, "ppt/_rels/presentation.xml.rels", presRels);
                    WriteZipBytes(zout, "[Content_Types].xml", types);
                }

                if (!File.Exists(slimPath) || new FileInfo(slimPath).Length < 64)
                {
                    error = "瘦包未写出";
                    TryDeleteFile(slimPath);
                    slimPath = null;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                TryDeleteFile(slimPath);
                slimPath = null;
                return false;
            }
        }

        private static bool TryApplyReplacements(
            string path,
            IDictionary<string, byte[]> replacements,
            out string error)
        {
            error = null;
            if (string.IsNullOrEmpty(path) || replacements == null || replacements.Count == 0)
            {
                return true;
            }

            try
            {
                using (ZipArchive zip = ZipFile.Open(path, ZipArchiveMode.Update))
                {
                    foreach (KeyValuePair<string, byte[]> kv in replacements)
                    {
                        if (kv.Value == null)
                        {
                            continue;
                        }

                        string name = kv.Key;
                        ZipArchiveEntry old = null;
                        foreach (ZipArchiveEntry e in zip.Entries)
                        {
                            if (NormPart(e.FullName).Equals(NormPart(name), StringComparison.OrdinalIgnoreCase))
                            {
                                old = e;
                                break;
                            }
                        }

                        if (old != null)
                        {
                            name = old.FullName;
                            old.Delete();
                        }

                        WriteZipBytes(zip, name, kv.Value);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static bool TryResolveSlidePart(
            Dictionary<string, ZipArchiveEntry> map,
            IDictionary<string, byte[]> replacements,
            int slideNo,
            out string slidePart,
            out string error)
        {
            slidePart = null;
            error = null;
            byte[] pres = ReadPart(map, replacements, "ppt/presentation.xml");
            byte[] rels = ReadPart(map, replacements, "ppt/_rels/presentation.xml.rels");
            if (pres == null || rels == null)
            {
                error = "没有 presentation.xml";
                return false;
            }

            XDocument presDoc = XDocument.Parse(Encoding.UTF8.GetString(pres));
            XElement lst = presDoc.Root == null ? null : presDoc.Root.Element(PNs + "sldIdLst");
            List<XElement> ids = lst == null
                ? new List<XElement>()
                : lst.Elements(PNs + "sldId").ToList();
            if (slideNo > ids.Count)
            {
                error = "页号超出 " + slideNo + "/" + ids.Count;
                return false;
            }

            string rid = (string)ids[slideNo - 1].Attribute(RNs + "id");
            XDocument relDoc = XDocument.Parse(Encoding.UTF8.GetString(rels));
            foreach (XElement rel in relDoc.Root.Elements(RelNs + "Relationship"))
            {
                if (!string.Equals((string)rel.Attribute("Id"), rid, StringComparison.Ordinal))
                {
                    continue;
                }

                slidePart = ResolvePart("ppt/presentation.xml", (string)rel.Attribute("Target"));
                if (HasPart(map, replacements, slidePart))
                {
                    return true;
                }
            }

            string fallback = "ppt/slides/slide" + slideNo + ".xml";
            if (HasPart(map, replacements, fallback))
            {
                slidePart = fallback;
                return true;
            }

            error = "找不到 slide 部件";
            return false;
        }

        private static HashSet<string> CollectSlideThemeClosure(
            Dictionary<string, ZipArchiveEntry> map,
            IDictionary<string, byte[]> replacements,
            string slidePart)
        {
            var need = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var q = new Queue<string>();
            Action<string> add = p =>
            {
                p = NormPart(p);
                if (string.IsNullOrEmpty(p) || !need.Add(p))
                {
                    return;
                }

                q.Enqueue(p);
                string rels = RelsOf(p);
                if (HasPart(map, replacements, rels) && need.Add(rels))
                {
                    q.Enqueue(rels);
                }
            };

            add("[Content_Types].xml");
            add("_rels/.rels");
            add("ppt/presentation.xml");
            add("ppt/_rels/presentation.xml.rels");
            add(slidePart);
            foreach (string extra in new[]
            {
                "ppt/presProps.xml",
                "ppt/viewProps.xml",
                "ppt/tableStyles.xml",
                "docProps/core.xml",
                "docProps/app.xml"
            })
            {
                if (HasPart(map, replacements, extra))
                {
                    add(extra);
                }
            }

            while (q.Count > 0)
            {
                string part = q.Dequeue();
                if (!part.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                byte[] xml = ReadPart(map, replacements, part);
                if (xml == null)
                {
                    continue;
                }

                string owner = OwnerOfRels(part);
                bool fromPresRels = part.Equals("ppt/_rels/presentation.xml.rels", StringComparison.OrdinalIgnoreCase);
                foreach (string target in RelTargets(xml, owner))
                {
                    string n = NormPart(target);
                    if (IsOtherSlide(n, slidePart))
                    {
                        continue;
                    }

                    if (n.StartsWith("ppt/fonts/", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (fromPresRels
                        && (n.StartsWith("ppt/slideMasters/", StringComparison.OrdinalIgnoreCase)
                            || n.StartsWith("ppt/slideLayouts/", StringComparison.OrdinalIgnoreCase)
                            || n.IndexOf("handout", StringComparison.OrdinalIgnoreCase) >= 0
                            || n.IndexOf("notesMaster", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        continue;
                    }

                    add(n);
                }
            }

            return need;
        }

        private static byte[] RewritePresentationOneSlide(
            byte[] presXml,
            string slidePart,
            Dictionary<string, ZipArchiveEntry> map,
            IDictionary<string, byte[]> replacements)
        {
            if (presXml == null)
            {
                return null;
            }

            XDocument doc = XDocument.Parse(Encoding.UTF8.GetString(presXml));
            XElement lst = doc.Root == null ? null : doc.Root.Element(PNs + "sldIdLst");
            if (lst == null)
            {
                return presXml;
            }

            byte[] relBytes = ReadPart(map, replacements, "ppt/_rels/presentation.xml.rels");
            XDocument relDoc = relBytes == null ? null : XDocument.Parse(Encoding.UTF8.GetString(relBytes));
            string keepRid = null;
            if (relDoc != null && relDoc.Root != null)
            {
                foreach (XElement rel in relDoc.Root.Elements(RelNs + "Relationship"))
                {
                    string resolved = ResolvePart("ppt/presentation.xml", (string)rel.Attribute("Target"));
                    if (resolved.Equals(slidePart, StringComparison.OrdinalIgnoreCase))
                    {
                        keepRid = (string)rel.Attribute("Id");
                        break;
                    }
                }
            }

            foreach (XElement id in lst.Elements(PNs + "sldId").ToList())
            {
                string rid = (string)id.Attribute(RNs + "id");
                if (keepRid == null || !string.Equals(rid, keepRid, StringComparison.Ordinal))
                {
                    id.Remove();
                }
            }

            return Encoding.UTF8.GetBytes(XmlBytes(doc));
        }

        private static byte[] RewritePresentationRels(byte[] relsXml, HashSet<string> need)
        {
            if (relsXml == null)
            {
                return null;
            }

            XDocument doc = XDocument.Parse(Encoding.UTF8.GetString(relsXml));
            if (doc.Root == null)
            {
                return relsXml;
            }

            foreach (XElement rel in doc.Root.Elements(RelNs + "Relationship").ToList())
            {
                string resolved = ResolvePart("ppt/presentation.xml", (string)rel.Attribute("Target"));
                string type = (string)rel.Attribute("Type") ?? "";
                bool isSlide = type.IndexOf("/slide", StringComparison.OrdinalIgnoreCase) >= 0
                    && type.IndexOf("slideMaster", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("slideLayout", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("notesSlide", StringComparison.OrdinalIgnoreCase) < 0;
                if (isSlide)
                {
                    if (!need.Contains(resolved))
                    {
                        rel.Remove();
                    }

                    continue;
                }

                if (type.IndexOf("/font", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    rel.Remove();
                    continue;
                }

                if (!need.Contains(resolved)
                    && type.IndexOf("theme", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("slideMaster", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("presProps", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("viewProps", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("tableStyles", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    rel.Remove();
                }
            }

            return Encoding.UTF8.GetBytes(XmlBytes(doc));
        }

        private static byte[] RewriteContentTypes(byte[] typesXml, HashSet<string> need)
        {
            if (typesXml == null)
            {
                return Encoding.UTF8.GetBytes(
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<Types xmlns=\"" + CtNs.NamespaceName + "\"/>");
            }

            XDocument doc = XDocument.Parse(Encoding.UTF8.GetString(typesXml));
            if (doc.Root == null)
            {
                return typesXml;
            }

            foreach (XElement o in doc.Root.Elements(CtNs + "Override").ToList())
            {
                string pn = ((string)o.Attribute("PartName") ?? "").TrimStart('/');
                if (!need.Contains(pn)
                    && !pn.Equals("ppt/presentation.xml", StringComparison.OrdinalIgnoreCase))
                {
                    o.Remove();
                }
            }

            return Encoding.UTF8.GetBytes(XmlBytes(doc));
        }

        private static bool IsOtherSlide(string part, string keepSlide)
        {
            if (Regex.IsMatch(part, @"^ppt/slides/slide\d+\.xml$", RegexOptions.IgnoreCase)
                && !part.Equals(keepSlide, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (Regex.IsMatch(part, @"^ppt/slides/_rels/slide\d+\.xml\.rels$", RegexOptions.IgnoreCase))
            {
                string owner = OwnerOfRels(part);
                return !owner.Equals(keepSlide, StringComparison.OrdinalIgnoreCase);
            }

            if (part.IndexOf("notesSlide", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static IEnumerable<string> RelTargets(byte[] relsXml, string fromPart)
        {
            XDocument doc;
            try
            {
                doc = XDocument.Parse(Encoding.UTF8.GetString(relsXml));
            }
            catch (Exception)
            {
                yield break;
            }

            if (doc.Root == null)
            {
                yield break;
            }

            foreach (XElement rel in doc.Root.Elements(RelNs + "Relationship"))
            {
                string target = (string)rel.Attribute("Target");
                string mode = (string)rel.Attribute("TargetMode");
                if (string.IsNullOrEmpty(target)
                    || string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                yield return ResolvePart(fromPart, target);
            }
        }

        private static string NormPart(string p)
        {
            return (p ?? "").Replace('\\', '/').TrimStart('/');
        }

        private static string DirOf(string part)
        {
            int i = part.LastIndexOf('/');
            return i < 0 ? "" : part.Substring(0, i + 1);
        }

        private static string ResolvePart(string fromPart, string target)
        {
            target = (target ?? "").Replace('\\', '/');
            if (target.StartsWith("/"))
            {
                return target.TrimStart('/');
            }

            string dir = DirOf(fromPart);
            var parts = new List<string>();
            foreach (string bit in (dir + target).Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (bit == ".")
                {
                    continue;
                }

                if (bit == "..")
                {
                    if (parts.Count > 0)
                    {
                        parts.RemoveAt(parts.Count - 1);
                    }
                }
                else
                {
                    parts.Add(bit);
                }
            }

            return string.Join("/", parts);
        }

        private static string RelsOf(string part)
        {
            int i = part.LastIndexOf('/');
            if (i < 0)
            {
                return "_rels/" + part + ".rels";
            }

            return part.Substring(0, i + 1) + "_rels/" + part.Substring(i + 1) + ".rels";
        }

        private static string OwnerOfRels(string relsPart)
        {
            int r = relsPart.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase);
            if (r >= 0)
            {
                string owner = relsPart.Substring(0, r + 1) + Path.GetFileName(relsPart);
                if (owner.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                {
                    owner = owner.Substring(0, owner.Length - 5);
                }

                return owner;
            }

            if (relsPart.StartsWith("_rels/", StringComparison.OrdinalIgnoreCase) && relsPart.EndsWith(".rels"))
            {
                return relsPart.Substring(6, relsPart.Length - 6 - 5);
            }

            return relsPart;
        }

        private static byte[] LookupReplacement(IDictionary<string, byte[]> replacements, string part)
        {
            if (replacements == null || string.IsNullOrEmpty(part))
            {
                return null;
            }

            byte[] data;
            if (replacements.TryGetValue(NormPart(part), out data))
            {
                return data;
            }

            return null;
        }

        private static bool HasPart(
            Dictionary<string, ZipArchiveEntry> map,
            IDictionary<string, byte[]> replacements,
            string part)
        {
            part = NormPart(part);
            return LookupReplacement(replacements, part) != null
                || (map != null && map.ContainsKey(part));
        }

        private static byte[] ReadPart(
            Dictionary<string, ZipArchiveEntry> map,
            IDictionary<string, byte[]> replacements,
            string part)
        {
            byte[] overlay = LookupReplacement(replacements, part);
            if (overlay != null)
            {
                return overlay;
            }

            ZipArchiveEntry e;
            if (map == null || !map.TryGetValue(NormPart(part), out e))
            {
                return null;
            }

            using (var ms = new MemoryStream())
            {
                using (Stream s = e.Open())
                {
                    s.CopyTo(ms);
                }

                return ms.ToArray();
            }
        }

        private static void CopyZipEntry(ZipArchiveEntry src, ZipArchive zout, string name)
        {
            ZipArchiveEntry dst = zout.CreateEntry(name, CompressionLevel.Fastest);
            using (Stream a = src.Open())
            using (Stream b = dst.Open())
            {
                a.CopyTo(b);
            }
        }

        private static void WriteZipBytes(ZipArchive zout, string name, byte[] data)
        {
            ZipArchiveEntry e = zout.CreateEntry(name, CompressionLevel.Fastest);
            using (Stream s = e.Open())
            {
                s.Write(data, 0, data.Length);
            }
        }

        private static string XmlBytes(XDocument doc)
        {
            string decl = doc.Declaration == null
                ? "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                : doc.Declaration.ToString();
            return decl + doc.ToString(SaveOptions.DisableFormatting);
        }
    }
}
