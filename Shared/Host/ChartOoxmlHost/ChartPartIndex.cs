using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Globalization;
using System.Xml.Linq;

namespace WordAddIn1.ChartOoxmlHost
{
    internal sealed class ChartPartRef
    {
        public int Ordinal { get; set; }
        public string RelId { get; set; }
        public string ChartPartName { get; set; }
        public XDocument ChartXml { get; set; }
        public string EmbedPartName { get; set; }
        public byte[] EmbedBytes { get; set; }
        public long ExtentCx { get; set; }
        public long ExtentCy { get; set; }
    }

    /// <summary>document.xml 正文 chart drawing 顺序 → rels → charts/chartN.xml。</summary>
    internal static class ChartPartIndex
    {
        public static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        public static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        public static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        public static readonly XNamespace Wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static bool TryIndexWordFlatOpc(string wordOpenXml, out List<ChartPartRef> parts, out string error)
        {
            parts = new List<ChartPartRef>();
            error = null;
            if (string.IsNullOrEmpty(wordOpenXml))
            {
                error = "WordOpenXML 空";
                return false;
            }

            XDocument pkg;
            try
            {
                pkg = XDocument.Parse(wordOpenXml);
            }
            catch (Exception ex)
            {
                error = "WordOpenXML 不是包: " + ex.Message;
                return false;
            }

            var byName = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement part in pkg.Descendants(Pkg + "part"))
            {
                string name = NormalizePartName(part.Attribute(Pkg + "name")?.Value);
                if (!string.IsNullOrEmpty(name))
                {
                    byName[name] = part;
                }
            }

            if (!byName.TryGetValue("/word/document.xml", out XElement docPart))
            {
                error = "包里没有 /word/document.xml";
                return false;
            }

            XDocument document = LoadXmlPart(docPart);
            XDocument rels = null;
            if (byName.TryGetValue("/word/_rels/document.xml.rels", out XElement relsPart))
            {
                rels = LoadXmlPart(relsPart);
            }

            if (document == null || document.Root == null)
            {
                error = "document.xml 空";
                return false;
            }

            Dictionary<string, string> chartRels = MapChartRels(rels);
            XElement body = document.Descendants(W + "body").FirstOrDefault();
            if (body == null)
            {
                error = "document.xml 无 body";
                return false;
            }

            int ordinal = 0;
            foreach (XElement drawing in body.Descendants(W + "drawing"))
            {
                if (IsInsideTextBox(drawing))
                {
                    continue;
                }

                XElement chartEl = drawing.Descendants(C + "chart").FirstOrDefault();
                if (chartEl == null)
                {
                    continue;
                }

                string rid = chartEl.Attribute(R + "id")?.Value;
                if (string.IsNullOrEmpty(rid) || !chartRels.TryGetValue(rid, out string target))
                {
                    error = "第 " + ordinal + " 张图对不上 rels";
                    parts = null;
                    return false;
                }

                string chartName = ResolveWordPart("/word/", target);
                if (!byName.TryGetValue(chartName, out XElement chartPart) || LoadXmlPart(chartPart) == null)
                {
                    error = "缺 chart 部件 " + chartName;
                    parts = null;
                    return false;
                }

                ReadExtent(drawing, out long cx, out long cy);
                string embedName = null;
                byte[] embedBytes = null;
                TryReadEmbed(byName, chartName, out embedName, out embedBytes);

                parts.Add(new ChartPartRef
                {
                    Ordinal = ordinal,
                    RelId = rid,
                    ChartPartName = chartName,
                    ChartXml = LoadXmlPart(chartPart),
                    EmbedPartName = embedName,
                    EmbedBytes = embedBytes,
                    ExtentCx = cx,
                    ExtentCy = cy
                });
                ordinal++;
            }

            return true;
        }

        public static bool TryApplyReplacementsToZip(string path, IDictionary<string, byte[]> replacements, out string error)
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
                    foreach (KeyValuePair<string, byte[]> pair in replacements)
                    {
                        string name = (pair.Key ?? "").TrimStart('/').Replace('\\', '/');
                        ZipArchiveEntry existing = zip.Entries.FirstOrDefault(e =>
                            string.Equals(e.FullName.Replace('\\', '/'), name, StringComparison.OrdinalIgnoreCase));
                        if (existing != null)
                        {
                            existing.Delete();
                        }

                        ZipArchiveEntry next = zip.CreateEntry(name, CompressionLevel.Optimal);
                        using (Stream s = next.Open())
                        {
                            s.Write(pair.Value, 0, pair.Value.Length);
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "回写 zip 失败: " + ex.Message;
                return false;
            }
        }

        public static string NormalizePartName(string name)
        {
            if (string.IsNullOrEmpty(name))
            {
                return "";
            }

            string t = name.Replace('\\', '/');
            if (!t.StartsWith("/", StringComparison.Ordinal))
            {
                t = "/" + t;
            }

            return t;
        }

        private static Dictionary<string, string> MapChartRels(XDocument rels)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            if (rels == null)
            {
                return map;
            }

            foreach (XElement rel in rels.Descendants())
            {
                if (rel.Name.LocalName != "Relationship")
                {
                    continue;
                }

                string type = rel.Attribute("Type")?.Value ?? "";
                if (type.IndexOf("/chart", StringComparison.OrdinalIgnoreCase) < 0
                    || type.IndexOf("chartColor", StringComparison.OrdinalIgnoreCase) >= 0
                    || type.IndexOf("chartStyle", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    continue;
                }

                string id = rel.Attribute("Id")?.Value;
                string target = rel.Attribute("Target")?.Value;
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                {
                    map[id] = target;
                }
            }

            return map;
        }

        private static string ResolveWordPart(string baseDir, string target)
        {
            string t = (target ?? "").Replace('\\', '/');
            if (t.StartsWith("/", StringComparison.Ordinal))
            {
                return NormalizePartName(t);
            }

            return NormalizePartName(baseDir.TrimEnd('/') + "/" + t);
        }

        private static void TryReadEmbed(
            Dictionary<string, XElement> byName,
            string chartName,
            out string embedName,
            out byte[] embedBytes)
        {
            embedName = null;
            embedBytes = null;
            string dir = chartName.Substring(0, chartName.LastIndexOf('/') + 1);
            string file = chartName.Substring(chartName.LastIndexOf('/') + 1);
            string relsName = dir + "_rels/" + file + ".rels";
            if (!byName.TryGetValue(relsName, out XElement relsPart))
            {
                return;
            }

            XDocument rels = LoadXmlPart(relsPart);
            if (rels == null)
            {
                return;
            }

            foreach (XElement rel in rels.Descendants())
            {
                if (rel.Name.LocalName != "Relationship")
                {
                    continue;
                }

                string type = rel.Attribute("Type")?.Value ?? "";
                if (type.IndexOf("package", StringComparison.OrdinalIgnoreCase) < 0
                    && type.IndexOf("oleObject", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    continue;
                }

                string target = rel.Attribute("Target")?.Value;
                if (string.IsNullOrEmpty(target))
                {
                    continue;
                }

                string partName = ResolveWordPart(dir, target);
                if (!byName.TryGetValue(partName, out XElement embedPart))
                {
                    continue;
                }

                XElement bin = embedPart.Element(Pkg + "binaryData");
                if (bin == null || string.IsNullOrEmpty(bin.Value))
                {
                    continue;
                }

                try
                {
                    embedName = partName;
                    embedBytes = Convert.FromBase64String(bin.Value.Trim());
                    return;
                }
                catch
                {
                }
            }
        }

        private static XDocument LoadXmlPart(XElement part)
        {
            if (part == null)
            {
                return null;
            }

            XElement xmlData = part.Element(Pkg + "xmlData");
            if (xmlData == null)
            {
                return null;
            }

            XElement root = xmlData.Elements().FirstOrDefault();
            if (root == null)
            {
                return null;
            }

            return new XDocument(new XElement(root));
        }

        private static bool IsInsideTextBox(XElement element)
        {
            return element != null && element.Ancestors(W + "txbxContent").Any();
        }

        private static void ReadExtent(XElement drawing, out long cx, out long cy)
        {
            cx = 0;
            cy = 0;
            XElement extent = drawing.Descendants(Wp + "extent").FirstOrDefault();
            if (extent == null)
            {
                return;
            }

            long.TryParse(extent.Attribute("cx")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out cx);
            long.TryParse(extent.Attribute("cy")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out cy);
        }
    }
}
