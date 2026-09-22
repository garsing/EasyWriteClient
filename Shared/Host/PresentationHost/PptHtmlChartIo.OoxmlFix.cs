using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using WordAddIn1.OpenFiles;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// WPP 建图后 SERIES 锁在 AddChart2 默认 4 格，COM 改 Formula 会 E_FAIL。
    /// 兜底：SaveCopyAs → 改 chart XML 的 c:f / c:ptCount / c:pt → 打开副本把图 Copy 回原页。
    /// 只在 WPP 且系列点数 ≠ 稿行数时走；PPT 仍走 ListObject Resize。
    /// </summary>
    internal static partial class PptHtmlChartIo
    {
        private static readonly XNamespace CNs =
            "http://schemas.openxmlformats.org/drawingml/2006/chart";

        private static readonly XNamespace PNs =
            "http://schemas.openxmlformats.org/presentationml/2006/main";

        private static readonly XNamespace RNs =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private static readonly XNamespace RelPkgNs =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        private static bool TryFixSeriesViaOoxml(
            object shape,
            PptHtmlChartGrid grid,
            List<string> warnings,
            out object newShape)
        {
            newShape = null;
            if (shape == null || grid == null || !grid.IsPourable)
            {
                return false;
            }

            object slide = null;
            object pres = null;
            object app = null;
            object copyPres = null;
            object copyChartShape = null;
            string tempPath = null;
            try
            {
                PourLog(warnings, "OOXML 修点：进入");
                slide = TryGetShapeSlide(shape);
                pres = TryGetShapePresentation(shape);
                if (slide == null || pres == null)
                {
                    PourLog(warnings, "OOXML 修点：取不到 slide/presentation");
                    return false;
                }

                PourLog(warnings, "OOXML 修点：slide/pres ok");
                app = WppCom.GetProperty(pres, "Application");
                string appName = TryPropString(app, "Name") ?? "";
                PourLog(warnings, "OOXML 修点：宿主 Name=" + appName);
                if (!LooksLikeWppApp(app))
                {
                    PourLog(warnings, "OOXML 修点：非 WPP 宿主（" + appName + "），跳过");
                    return false;
                }

                PourLog(warnings, "OOXML 修点：WPP 宿主确认");

                int shapeId = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
                int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
                PourLog(warnings, "OOXML 修点：shapeId=" + shapeId + " slideIndex=" + slideIndex);
                if (shapeId < 1 || slideIndex < 1)
                {
                    PourLog(warnings, "OOXML 修点：shapeId/slideIndex 非法 " + shapeId + "/" + slideIndex);
                    return false;
                }

                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wpp-chart-" + Guid.NewGuid().ToString("N") + ".pptx");
                PourLog(warnings, "OOXML 修点：SaveCopyAs → " + tempPath);
                try
                {
                    WppCom.Invoke(pres, "SaveCopyAs", tempPath);
                }
                catch (Exception ex)
                {
                    PourLog(warnings, "SaveCopyAs(1参) 失败: " + FormatComError(ex));
                    try
                    {
                        WppCom.Invoke(pres, "SaveCopyAs", tempPath, 24);
                    }
                    catch (Exception ex2)
                    {
                        PourLog(warnings, "SaveCopyAs(2参) 失败: " + FormatComError(ex2));
                        return false;
                    }
                }

                if (!File.Exists(tempPath))
                {
                    PourLog(warnings, "OOXML 修点：SaveCopyAs 未落盘 " + tempPath);
                    return false;
                }

                PourLog(warnings, "OOXML 修点：副本已落盘 " + tempPath + " size=" + new FileInfo(tempPath).Length);

                if (!TryRewriteChartXmlInPackage(tempPath, slideIndex, shapeId, grid, warnings))
                {
                    return false;
                }

                PourLog(warnings, "OOXML 修点：打开副本 " + tempPath);
                object presentations = WppCom.GetProperty(app, "Presentations");
                copyPres = TryOpenCopyPresentation(presentations, tempPath, warnings);
                if (copyPres == null)
                {
                    PourLog(warnings, "OOXML 修点：打开副本失败");
                    return false;
                }

                PourLog(warnings, "OOXML 修点：副本已打开");

                object copySlides = WppCom.GetProperty(copyPres, "Slides");
                object copySlide = WppCom.GetIndexed(copySlides, slideIndex);
                object copyShapes = WppCom.GetProperty(copySlide, "Shapes");
                int copyCount = Convert.ToInt32(WppCom.GetProperty(copyShapes, "Count") ?? 0);
                for (int i = 1; i <= copyCount; i++)
                {
                    object s = WppCom.GetIndexed(copyShapes, i);
                    if (Convert.ToInt32(WppCom.GetProperty(s, "Id") ?? 0) == shapeId)
                    {
                        copyChartShape = s;
                        break;
                    }
                }

                if (copyChartShape == null)
                {
                    PourLog(warnings, "OOXML 修点：副本里找不到 shapeId=" + shapeId);
                    return false;
                }

                int copyPts = ReadSeriesRowCount(TryGetChart(copyChartShape));
                PourLog(warnings, "OOXML 修点：副本图点数=" + copyPts + "（稿要 " + grid.Rows.Count + "）");
                if (copyPts != grid.Rows.Count)
                {
                    return false;
                }

                PourLog(warnings, "OOXML 修点：Copy 副本图");
                WppCom.Invoke(copyChartShape, "Copy");
                object shapes = WppCom.GetProperty(slide, "Shapes");
                object pasted = TryPasteChart(shapes, warnings);
                if (pasted == null)
                {
                    PourLog(warnings, "OOXML 修点：粘贴回原页失败");
                    return false;
                }

                PourLog(warnings, "OOXML 修点：已粘贴回原页");

                TryCopyBox(shape, pasted);
                object pastedChart = TryGetChart(pasted);
                int livePts = ReadSeriesRowCount(pastedChart);
                PourLog(warnings, "OOXML 修点：贴回后点数=" + livePts);
                if (livePts != grid.Rows.Count)
                {
                    TryDelete(pasted);
                    return false;
                }

                TryDelete(shape);
                newShape = pasted;
                PourLog(warnings, "OOXML 修点完成：删旧图，新图 " + livePts + " 点");
                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML 修点失败: " + FormatComError(ex));
                return false;
            }
            finally
            {
                if (copyPres != null)
                {
                    try
                    {
                        WppCom.Invoke(copyPres, "Close");
                    }
                    catch (Exception)
                    {
                    }
                }

                if (!string.IsNullOrEmpty(tempPath))
                {
                    try
                    {
                        File.Delete(tempPath);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        /// <summary>
        /// WPP 内嵌簿由 Excel 12.0 渲染，chart.Application.Name 会报「Microsoft PowerPoint」。
        /// 不能只看 Name：试读 WPP 专有属性，再退化到 Name 含 wps/wpp。
        /// </summary>
        private static bool LooksLikeWppApp(object app)
        {
            if (app == null)
            {
                return false;
            }

            try
            {
                object v = WppCom.GetProperty(app, "WpsPresentation");
                if (v != null)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                object v = WppCom.GetProperty(app, "ProductCode");
                if (v != null)
                {
                    return true;
                }
            }
            catch (Exception)
            {
            }

            string name = TryPropString(app, "Name") ?? "";
            return name.IndexOf("wps", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("wpp", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static object TryOpenCopyPresentation(object presentations, string fullPath, List<string> warnings)
        {
            if (presentations == null || string.IsNullOrEmpty(fullPath))
            {
                return null;
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", fullPath);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "Open(仅路径) 失败: " + FormatComError(ex));
            }

            try
            {
                object opened = WppCom.Invoke(presentations, "Open", fullPath, false, false, false);
                if (opened != null)
                {
                    return opened;
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "Open(4参) 失败: " + FormatComError(ex));
            }

            try
            {
                return WppCom.Invoke(presentations, "Open", fullPath, true);
            }
            catch (Exception ex)
            {
                PourLog(warnings, "Open(WithWindow) 失败: " + FormatComError(ex));
                return null;
            }
        }

        private static object TryPasteChart(object shapes, List<string> warnings)
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
                    object first = TryGetIndexedSafe(pasted, 1) ?? pasted;
                    return first;
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "Paste 失败: " + FormatComError(ex));
            }

            try
            {
                object pasted = WppCom.Invoke(shapes, "PasteSpecial", 0);
                return TryGetIndexedSafe(pasted, 1) ?? pasted;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "PasteSpecial 失败: " + FormatComError(ex));
                return null;
            }
        }

        private static object TryGetIndexedSafe(object collection, int index)
        {
            try
            {
                return WppCom.GetIndexed(collection, index);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void TryCopyBox(object from, object to)
        {
            if (from == null || to == null)
            {
                return;
            }

            foreach (string prop in new[] { "Left", "Top", "Width", "Height" })
            {
                try
                {
                    object v = WppCom.GetProperty(from, prop);
                    if (v != null)
                    {
                        WppCom.TrySetProperty(to, prop, v);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static object TryGetChartShape(object chart)
        {
            try
            {
                return WppCom.GetProperty(chart, "Parent");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetShapeSlide(object shape)
        {
            try
            {
                object parent = WppCom.GetProperty(shape, "Parent");
                if (parent == null)
                {
                    return null;
                }

                object idx = null;
                try
                {
                    idx = WppCom.GetProperty(parent, "SlideIndex");
                }
                catch (Exception)
                {
                }

                if (idx != null)
                {
                    return parent;
                }

                object grand = WppCom.GetProperty(parent, "Parent");
                return grand;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static object TryGetShapePresentation(object shape)
        {
            try
            {
                object slide = TryGetShapeSlide(shape);
                return slide == null ? null : WppCom.GetProperty(slide, "Parent");
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool TryRewriteChartXmlInPackage(
            string pptxPath,
            int slideIndex,
            int shapeId,
            PptHtmlChartGrid grid,
            List<string> warnings)
        {
            string slidePart = "ppt/slides/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
            string slideRels = "ppt/slides/_rels/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml.rels";
            string chartPart = null;
            string tempOut = pptxPath + ".fix";
            try
            {
                using (ZipArchive zip = ZipFile.Open(pptxPath, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry slideEntry = FindEntry(zip, slidePart);
                    ZipArchiveEntry relsEntry = FindEntry(zip, slideRels);
                    if (slideEntry == null || relsEntry == null)
                    {
                        PourLog(warnings, "OOXML 修点：缺 slide 部件 " + slidePart);
                        return false;
                    }

                    XDocument slideDoc;
                    using (Stream s = slideEntry.Open())
                    {
                        slideDoc = XDocument.Load(s);
                    }

                    string relId = FindChartRelId(slideDoc, shapeId);
                    if (string.IsNullOrEmpty(relId))
                    {
                        PourLog(warnings, "OOXML 修点：slide 里找不到 shapeId=" + shapeId + " 的 chart 关系");
                        return false;
                    }

                    XDocument relsDoc;
                    using (Stream s = relsEntry.Open())
                    {
                        relsDoc = XDocument.Load(s);
                    }

                    string target = FindRelTarget(relsDoc, relId);
                    if (string.IsNullOrEmpty(target))
                    {
                        PourLog(warnings, "OOXML 修点：rels 里找不到 " + relId);
                        return false;
                    }

                    chartPart = NormalizeChartPartPath(target);
                    ZipArchiveEntry chartEntry = FindEntry(zip, chartPart);
                    if (chartEntry == null)
                    {
                        PourLog(warnings, "OOXML 修点：缺 chart 部件 " + chartPart);
                        return false;
                    }

                    XDocument chartDoc;
                    using (Stream s = chartEntry.Open())
                    {
                        chartDoc = XDocument.Load(s);
                    }

                    if (!RewriteChartSeries(chartDoc, grid, warnings))
                    {
                        return false;
                    }

                    chartEntry.Delete();
                    ZipArchiveEntry fresh = zip.CreateEntry(chartEntry.FullName);
                    using (Stream s = fresh.Open())
                    {
                        chartDoc.Save(s);
                    }
                }

                PourLog(warnings, "OOXML 修点：已改 " + chartPart + " → " + grid.Rows.Count + " 点");
                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML 修点改包失败: " + FormatComError(ex));
                return false;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempOut))
                    {
                        File.Delete(tempOut);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static string FindChartRelId(XDocument slideDoc, int shapeId)
        {
            foreach (XElement frame in slideDoc.Descendants(PNs + "graphicFrame"))
            {
                XElement cNvPr = frame.Descendants(PNs + "cNvPr").FirstOrDefault();
                if (cNvPr == null)
                {
                    continue;
                }

                string idAttr = (string)cNvPr.Attribute("id");
                if (!string.Equals(idAttr, shapeId.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
                {
                    continue;
                }

                XElement chartEl = frame.Descendants(CNs + "chart").FirstOrDefault();
                if (chartEl == null)
                {
                    continue;
                }

                return (string)chartEl.Attribute(RNs + "id");
            }

            return null;
        }

        private static string FindRelTarget(XDocument relsDoc, string relId)
        {
            foreach (XElement rel in relsDoc.Descendants(RelPkgNs + "Relationship"))
            {
                if (string.Equals((string)rel.Attribute("Id"), relId, StringComparison.Ordinal))
                {
                    return (string)rel.Attribute("Target");
                }
            }

            return null;
        }

        private static string NormalizeChartPartPath(string target)
        {
            if (string.IsNullOrEmpty(target))
            {
                return null;
            }

            string t = target.Replace('\\', '/');
            if (t.StartsWith("../", StringComparison.Ordinal))
            {
                t = "ppt/" + t.Substring(3);
            }
            else if (!t.StartsWith("ppt/", StringComparison.Ordinal))
            {
                t = "ppt/" + t.TrimStart('/');
            }

            return t;
        }

        private static bool RewriteChartSeries(XDocument chartDoc, PptHtmlChartGrid grid, List<string> warnings)
        {
            int rows = grid.Rows.Count;
            int last = rows + 1;
            string lastText = last.ToString(CultureInfo.InvariantCulture);
            var cats = new List<string>();
            foreach (List<string> row in grid.Rows)
            {
                cats.Add(row != null && row.Count > 0 ? (row[0] ?? "") : "");
            }

            List<XElement> series = chartDoc.Descendants(CNs + "ser").ToList();
            if (series.Count == 0)
            {
                PourLog(warnings, "OOXML 修点：chart 无 c:ser");
                return false;
            }

            int si = 0;
            for (int c = 0; c < grid.Columns.Count; c++)
            {
                if (grid.Columns[c].Role == "category")
                {
                    continue;
                }

                if (si >= series.Count)
                {
                    break;
                }

                XElement ser = series[si];
                si++;
                string col = ColLetter(c + 1);
                string nameRef = "Sheet1!$" + col + "$1";
                string catsRef = "Sheet1!$A$2:$A$" + lastText;
                string valsRef = "Sheet1!$" + col + "$2:$" + col + "$" + lastText;

                var values = new List<string>();
                foreach (List<string> row in grid.Rows)
                {
                    values.Add(row != null && c < row.Count ? (row[c] ?? "") : "");
                }

                RewriteSerRef(ser, CNs + "tx", CNs + "strRef", nameRef,
                    new List<string> { grid.Columns[c].Name ?? "" });
                RewriteSerRef(ser, CNs + "cat", CNs + "strRef", catsRef, cats);
                RewriteSerRef(ser, CNs + "cat", CNs + "numRef", catsRef, cats);
                RewriteSerRef(ser, CNs + "val", CNs + "numRef", valsRef, values);
                RewriteSerRef(ser, CNs + "val", CNs + "strRef", valsRef, values);
            }

            return true;
        }

        private static void RewriteSerRef(
            XElement ser,
            XName parentName,
            XName refName,
            string formula,
            List<string> points)
        {
            XElement parent = ser.Element(parentName);
            if (parent == null)
            {
                return;
            }

            XElement refEl = parent.Element(refName);
            if (refEl == null)
            {
                return;
            }

            XElement f = refEl.Element(CNs + "f");
            if (f != null)
            {
                f.Value = formula;
            }

            XElement cache = refEl.Element(CNs + "strCache") ?? refEl.Element(CNs + "numCache");
            if (cache == null)
            {
                return;
            }

            XElement ptCount = cache.Element(CNs + "ptCount");
            if (ptCount != null)
            {
                ptCount.SetAttributeValue("val", points.Count.ToString(CultureInfo.InvariantCulture));
            }

            foreach (XElement pt in cache.Elements(CNs + "pt").ToList())
            {
                pt.Remove();
            }

            for (int i = 0; i < points.Count; i++)
            {
                var pt = new XElement(CNs + "pt");
                pt.SetAttributeValue("idx", i.ToString(CultureInfo.InvariantCulture));
                pt.Add(new XElement(CNs + "v", points[i] ?? ""));
                cache.Add(pt);
            }
        }

        private static ZipArchiveEntry FindEntry(ZipArchive zip, string fullName)
        {
            ZipArchiveEntry e = zip.GetEntry(fullName);
            if (e != null)
            {
                return e;
            }

            return zip.Entries.FirstOrDefault(x =>
                string.Equals(x.FullName.Replace('\\', '/'), fullName, StringComparison.OrdinalIgnoreCase));
        }
    }
}
