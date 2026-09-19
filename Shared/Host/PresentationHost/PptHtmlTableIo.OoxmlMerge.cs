using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 读合并拓扑：SaveCopyAs 后解包 DrawingML（a:tbl）。
    /// 文字/皮/列宽等仍走 Interop；OOXML 失败则回退 Shape 旁证。
    /// </summary>
    internal static partial class PptHtmlTableIo
    {
        private static readonly XNamespace ANs =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

        private static readonly XNamespace PNs =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

        private static readonly XNamespace RNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static readonly XNamespace RelPkgNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private static readonly object OoxmlCacheLock = new object();
        private static string OoxmlCacheKey;
        private static Dictionary<string, MergeMap> OoxmlCacheMaps;
        private static string OoxmlCacheTempPath;
        private static DateTime OoxmlCacheUtc;

        private static MergeMap ResolveMergeMap(
            PowerPoint.Shape shape,
            PowerPoint.Table table,
            int useRows,
            int useCols,
            List<string> warnings)
        {
            if (TryGetMergeMapFromOoxml(shape, useRows, useCols, out MergeMap fromOoxml, warnings)
                && fromOoxml != null)
            {
                return fromOoxml;
            }

            return BuildMergeMap(table, useRows, useCols, warnings);
        }

        private static MergeMap ResolveMergeMapWpp(
            object shape,
            object table,
            int useRows,
            int useCols,
            List<string> warnings)
        {
            if (TryGetMergeMapFromOoxmlWpp(shape, useRows, useCols, out MergeMap fromOoxml, warnings)
                && fromOoxml != null)
            {
                return fromOoxml;
            }

            return BuildMergeMapWpp(table, useRows, useCols, warnings);
        }

        private static bool TryGetMergeMapFromOoxml(
            PowerPoint.Shape shape,
            int useRows,
            int useCols,
            out MergeMap map,
            List<string> warnings)
        {
            map = null;
            try
            {
                if (!TryGetSlideAndPresentation(shape, out PowerPoint.Slide slide, out PowerPoint.Presentation pres))
                {
                    warnings.Add("OOXML 合并：无法取到 Slide/Presentation，回退 Shape 旁证");
                    return false;
                }

                int slideIndex = slide.SlideIndex;
                int shapeId = shape.Id;
                string cacheKey = BuildPresCacheKey(pres);
                if (!TryEnsureOoxmlCache(cacheKey, () => SaveCopyAsPowerPoint(pres), warnings))
                {
                    return false;
                }

                string mapKey = slideIndex.ToString(CultureInfo.InvariantCulture) + ":"
                    + shapeId.ToString(CultureInfo.InvariantCulture);
                if (!OoxmlCacheMaps.TryGetValue(mapKey, out map) || map == null)
                {
                    warnings.Add("OOXML 合并：未找到 shapeId=" + shapeId + " 的 a:tbl，回退 Shape 旁证");
                    return false;
                }

                if (!MergeMapFits(map, useRows, useCols))
                {
                    warnings.Add("OOXML 合并：行列与现表不一致，回退 Shape 旁证");
                    map = null;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                warnings.Add("OOXML 合并失败，回退 Shape 旁证: " + ex.Message);
                map = null;
                return false;
            }
        }

        private static bool TryGetMergeMapFromOoxmlWpp(
            object shape,
            int useRows,
            int useCols,
            out MergeMap map,
            List<string> warnings)
        {
            map = null;
            try
            {
                object slide = WppCom.GetProperty(shape, "Parent");
                object pres = slide == null ? null : WppCom.GetProperty(slide, "Parent");
                if (slide == null || pres == null)
                {
                    warnings.Add("OOXML 合并(WPP)：无 Slide/Presentation，回退旁证");
                    return false;
                }

                int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
                int shapeId = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
                if (slideIndex < 1 || shapeId <= 0)
                {
                    warnings.Add("OOXML 合并(WPP)：SlideIndex/Id 无效，回退旁证");
                    return false;
                }

                string cacheKey = BuildPresCacheKeyWpp(pres);
                if (!TryEnsureOoxmlCache(cacheKey, () => SaveCopyAsWpp(pres), warnings))
                {
                    return false;
                }

                string mapKey = slideIndex.ToString(CultureInfo.InvariantCulture) + ":"
                    + shapeId.ToString(CultureInfo.InvariantCulture);
                if (!OoxmlCacheMaps.TryGetValue(mapKey, out map) || map == null)
                {
                    warnings.Add("OOXML 合并(WPP)：未找到 a:tbl，回退旁证");
                    return false;
                }

                if (!MergeMapFits(map, useRows, useCols))
                {
                    warnings.Add("OOXML 合并(WPP)：行列不一致，回退旁证");
                    map = null;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                warnings.Add("OOXML 合并(WPP)失败，回退旁证: " + ex.Message);
                map = null;
                return false;
            }
        }

        private static bool TryEnsureOoxmlCache(
            string cacheKey,
            Func<string> saveCopyAs,
            List<string> warnings)
        {
            lock (OoxmlCacheLock)
            {
                if (OoxmlCacheMaps != null
                    && string.Equals(OoxmlCacheKey, cacheKey, StringComparison.Ordinal)
                    && (DateTime.UtcNow - OoxmlCacheUtc).TotalSeconds < 8
                    && !string.IsNullOrEmpty(OoxmlCacheTempPath)
                    && File.Exists(OoxmlCacheTempPath))
                {
                    return true;
                }

                ClearOoxmlCache_NoLock();
                string tempPath = saveCopyAs();
                if (string.IsNullOrEmpty(tempPath) || !File.Exists(tempPath))
                {
                    warnings.Add("OOXML 合并：SaveCopyAs 失败");
                    return false;
                }

                Dictionary<string, MergeMap> maps = ParseAllTableMergeMaps(tempPath, warnings);
                OoxmlCacheKey = cacheKey;
                OoxmlCacheMaps = maps ?? new Dictionary<string, MergeMap>(StringComparer.Ordinal);
                OoxmlCacheTempPath = tempPath;
                OoxmlCacheUtc = DateTime.UtcNow;
                return true;
            }
        }

        private static void ClearOoxmlCache_NoLock()
        {
            OoxmlCacheKey = null;
            OoxmlCacheMaps = null;
            OoxmlCacheUtc = DateTime.MinValue;
            if (string.IsNullOrEmpty(OoxmlCacheTempPath))
            {
                return;
            }

            try
            {
                if (File.Exists(OoxmlCacheTempPath))
                {
                    File.Delete(OoxmlCacheTempPath);
                }
            }
            catch
            {
            }

            OoxmlCacheTempPath = null;
        }

        private static string SaveCopyAsPowerPoint(PowerPoint.Presentation presentation)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-ppt-merge-" + Guid.NewGuid().ToString("N") + ".pptx");
            presentation.SaveCopyAs(path, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            return path;
        }

        private static string SaveCopyAsWpp(object presentation)
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ew-wpp-merge-" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                WppCom.Invoke(presentation, "SaveCopyAs", path);
            }
            catch
            {
                WppCom.Invoke(presentation, "SaveCopyAs", path, 24);
            }

            return path;
        }

        private static string BuildPresCacheKey(PowerPoint.Presentation presentation)
        {
            try
            {
                string full = presentation.FullName;
                int n = presentation.Slides.Count;
                if (!string.IsNullOrWhiteSpace(full))
                {
                    return "ppt:" + full + "@" + n;
                }

                return "ppt:" + presentation.Name + "@" + n + "@"
                    + System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(presentation);
            }
            catch
            {
                return "ppt:anon@" + Guid.NewGuid().ToString("N");
            }
        }

        private static string BuildPresCacheKeyWpp(object presentation)
        {
            try
            {
                string full = Convert.ToString(WppCom.GetProperty(presentation, "FullName") ?? "");
                object slides = WppCom.GetProperty(presentation, "Slides");
                int count = slides == null
                    ? 0
                    : Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                if (!string.IsNullOrWhiteSpace(full))
                {
                    return "wpp:" + full + "@" + count;
                }

                string name = Convert.ToString(WppCom.GetProperty(presentation, "Name") ?? "anon");
                return "wpp:" + name + "@" + count;
            }
            catch
            {
                return "wpp:anon@" + Guid.NewGuid().ToString("N");
            }
        }

        private static bool TryGetSlideAndPresentation(
            PowerPoint.Shape shape,
            out PowerPoint.Slide slide,
            out PowerPoint.Presentation presentation)
        {
            slide = null;
            presentation = null;
            try
            {
                object parent = shape.Parent;
                slide = parent as PowerPoint.Slide;
                if (slide == null)
                {
                    PowerPoint.Shapes shapes = parent as PowerPoint.Shapes;
                    if (shapes != null)
                    {
                        slide = shapes.Parent as PowerPoint.Slide;
                    }
                }

                if (slide == null)
                {
                    return false;
                }

                presentation = slide.Parent as PowerPoint.Presentation;
                return presentation != null;
            }
            catch
            {
                return false;
            }
        }

        private static Dictionary<string, MergeMap> ParseAllTableMergeMaps(
            string pptxPath,
            List<string> warnings)
        {
            var result = new Dictionary<string, MergeMap>(StringComparer.Ordinal);
            using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
            {
                List<string> slideParts = ListSlidePartPaths(zip);
                for (int i = 0; i < slideParts.Count; i++)
                {
                    int slideIndex = i + 1;
                    ZipArchiveEntry entry = zip.GetEntry(slideParts[i]);
                    if (entry == null)
                    {
                        // Zip 内路径大小写/分隔符差异
                        entry = zip.Entries.FirstOrDefault(e =>
                            string.Equals(
                                e.FullName.Replace('\\', '/'),
                                slideParts[i],
                                StringComparison.OrdinalIgnoreCase));
                    }

                    if (entry == null)
                    {
                        continue;
                    }

                    XDocument doc;
                    using (Stream s = entry.Open())
                    {
                        doc = XDocument.Load(s);
                    }

                    foreach (XElement frame in doc.Descendants(PNs + "graphicFrame"))
                    {
                        XElement cNvPr = frame.Descendants(PNs + "cNvPr").FirstOrDefault();
                        if (cNvPr == null)
                        {
                            continue;
                        }

                        string idAttr = (string)cNvPr.Attribute("id");
                        if (string.IsNullOrEmpty(idAttr))
                        {
                            continue;
                        }

                        XElement tbl = frame.Descendants(ANs + "tbl").FirstOrDefault();
                        if (tbl == null)
                        {
                            continue;
                        }

                        MergeMap map = ParseDrawingMlTableMerges(tbl);
                        if (map == null)
                        {
                            continue;
                        }

                        string key = slideIndex.ToString(CultureInfo.InvariantCulture) + ":" + idAttr;
                        result[key] = map;
                    }
                }
            }

            if (result.Count == 0)
            {
                warnings.Add("OOXML 合并：副本中未解析到任何 a:tbl");
            }

            return result;
        }

        private static List<string> ListSlidePartPaths(ZipArchive zip)
        {
            ZipArchiveEntry pres = FindZipEntry(zip, "ppt/presentation.xml");
            ZipArchiveEntry rels = FindZipEntry(zip, "ppt/_rels/presentation.xml.rels");
            if (pres == null || rels == null)
            {
                return zip.Entries
                    .Select(e => e.FullName.Replace('\\', '/'))
                    .Where(p => p.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                        && p.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        && p.IndexOf("_rels", StringComparison.OrdinalIgnoreCase) < 0)
                    .OrderBy(p => p, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            XDocument relDoc;
            using (Stream s = rels.Open())
            {
                relDoc = XDocument.Load(s);
            }

            var ridToTarget = new Dictionary<string, string>(StringComparer.Ordinal);
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

            XDocument presDoc;
            using (Stream s = pres.Open())
            {
                presDoc = XDocument.Load(s);
            }

            var ordered = new List<string>();
            XElement sldIdLst = presDoc.Root.Element(PNs + "sldIdLst");
            if (sldIdLst == null)
            {
                return ridToTarget.Values
                    .Where(v => v.IndexOf("/slides/slide", StringComparison.OrdinalIgnoreCase) >= 0)
                    .OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            foreach (XElement sldId in sldIdLst.Elements(PNs + "sldId"))
            {
                string rid = (string)sldId.Attribute(RNs + "id");
                if (string.IsNullOrEmpty(rid))
                {
                    continue;
                }

                if (ridToTarget.TryGetValue(rid, out string path))
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

            return zip.Entries.FirstOrDefault(x =>
                string.Equals(x.FullName.Replace('\\', '/'), fullName, StringComparison.OrdinalIgnoreCase));
        }

        private static MergeMap ParseDrawingMlTableMerges(XElement tbl)
        {
            if (tbl == null)
            {
                return null;
            }

            List<XElement> trs = tbl.Elements(ANs + "tr").ToList();
            if (trs.Count == 0)
            {
                return new MergeMap();
            }

            int colCount = 0;
            var rawRows = new List<List<TcInfo>>();
            foreach (XElement tr in trs)
            {
                var cells = new List<TcInfo>();
                int slots = 0;
                foreach (XElement tc in tr.Elements(ANs + "tc"))
                {
                    var info = new TcInfo
                    {
                        GridSpan = ReadPositiveIntAttr(tc, "gridSpan", 1),
                        RowSpan = ReadPositiveIntAttr(tc, "rowSpan", 1),
                        HMerge = IsMergeAttr(tc, "hMerge"),
                        VMerge = IsMergeAttr(tc, "vMerge")
                    };
                    cells.Add(info);
                    slots += Math.Max(1, info.GridSpan);
                }

                if (slots > colCount)
                {
                    colCount = slots;
                }

                rawRows.Add(cells);
            }

            int rowCount = rawRows.Count;
            if (colCount < 1)
            {
                return new MergeMap();
            }

            var map = new MergeMap();
            var occupied = new bool[rowCount, colCount];

            for (int r = 0; r < rowCount; r++)
            {
                int cursor = 0;
                foreach (TcInfo tc in rawRows[r])
                {
                    while (cursor < colCount && occupied[r, cursor])
                    {
                        cursor++;
                    }

                    if (cursor >= colCount)
                    {
                        break;
                    }

                    if (tc.HMerge || tc.VMerge)
                    {
                        map.Covered.Add(Pack(r + 1, cursor + 1));
                        occupied[r, cursor] = true;
                        cursor += 1;
                        continue;
                    }

                    int rs = Math.Max(1, tc.RowSpan);
                    int cs = Math.Max(1, tc.GridSpan);
                    if (r + rs > rowCount)
                    {
                        rs = rowCount - r;
                    }

                    if (cursor + cs > colCount)
                    {
                        cs = colCount - cursor;
                    }

                    map.Anchors[Pack(r + 1, cursor + 1)] = new MergeSpan { RowSpan = rs, ColSpan = cs };
                    for (int i = 0; i < rs; i++)
                    {
                        for (int j = 0; j < cs; j++)
                        {
                            int rr = r + i;
                            int cc = cursor + j;
                            if (rr >= rowCount || cc >= colCount)
                            {
                                continue;
                            }

                            occupied[rr, cc] = true;
                            if (i != 0 || j != 0)
                            {
                                map.Covered.Add(Pack(rr + 1, cc + 1));
                            }
                        }
                    }

                    cursor += cs;
                }
            }

            for (int r = 1; r <= rowCount; r++)
            {
                for (int c = 1; c <= colCount; c++)
                {
                    int key = Pack(r, c);
                    if (!map.Covered.Contains(key) && !map.Anchors.ContainsKey(key))
                    {
                        map.Anchors[key] = new MergeSpan { RowSpan = 1, ColSpan = 1 };
                    }
                }
            }

            return map;
        }

        private static int ReadPositiveIntAttr(XElement el, string name, int fallback)
        {
            string raw = (string)el.Attribute(name);
            if (string.IsNullOrEmpty(raw))
            {
                return fallback;
            }

            if (int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int n) && n > 0)
            {
                return n;
            }

            return fallback;
        }

        private static bool IsMergeAttr(XElement el, string name)
        {
            XAttribute a = el.Attribute(name);
            if (a == null)
            {
                return false;
            }

            string v = ((string)a ?? "").Trim();
            if (v.Length == 0)
            {
                return true;
            }

            return v == "1"
                || v.Equals("true", StringComparison.OrdinalIgnoreCase)
                || v.Equals("on", StringComparison.OrdinalIgnoreCase);
        }

        private static bool MergeMapFits(MergeMap map, int useRows, int useCols)
        {
            if (map == null || useRows < 1 || useCols < 1)
            {
                return false;
            }

            int maxR = 0;
            int maxC = 0;
            foreach (KeyValuePair<int, MergeSpan> kv in map.Anchors)
            {
                Unpack(kv.Key, out int r, out int c);
                int er = r + Math.Max(1, kv.Value.RowSpan) - 1;
                int ec = c + Math.Max(1, kv.Value.ColSpan) - 1;
                if (er > maxR) maxR = er;
                if (ec > maxC) maxC = ec;
            }

            foreach (int key in map.Covered)
            {
                Unpack(key, out int r, out int c);
                if (r > maxR) maxR = r;
                if (c > maxC) maxC = c;
            }

            return maxR >= useRows && maxC >= useCols;
        }

        private sealed class TcInfo
        {
            public int GridSpan;
            public int RowSpan;
            public bool HMerge;
            public bool VMerge;
        }
    }
}
