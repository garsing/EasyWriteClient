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
    /// 写自定义 RGB 后同一套路剥 srgbClr 上的 tint/shade（默认 ChartStyle 会扣，COM 清不掉）。
    /// 渐变/标记色/线色 COM 写不稳、读常空：同一趟把 c:spPr / c:marker 钉进 XML。
    /// 只在 WPP 走；PPT 仍走 ListObject Resize / COM 写色。
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

        private static readonly XNamespace ANs =
            "http://schemas.openxmlformats.org/drawingml/2006/main";

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

            int want = grid.Rows.Count;
            return TryEditChartXmlAndCopyBack(
                shape,
                warnings,
                "修点",
                (doc, w) => RewriteChartSeries(doc, grid, w),
                copyShape => ReadSeriesRowCount(TryGetChart(copyShape)) == want,
                out newShape);
        }

        /// <summary>
        /// 默认 ChartStyle 会把 tint 扣在自定义 srgb 上，COM 清不掉。
        /// 渐变/标记/线 COM 写完也常读空。写色后 SaveCopyAs → 钉 c:spPr/c:marker 并剥 tint → Copy 回原页。
        /// </summary>
        private static void TryStripWppSrgbTintAfterChrome(
            ref object shape,
            ref object chart,
            ChartStyleSnap snap,
            List<string> warnings)
        {
            if (shape == null || !SnapNeedsWppOoxmlChrome(snap))
            {
                return;
            }

            DismissChartExcelUiForChart(chart);
            if (!TryEditChartXmlAndCopyBack(
                shape,
                warnings,
                "钉皮",
                (doc, w) =>
                {
                    bool pinned = RewriteSeriesChromeFromSnap(doc, snap, w);
                    bool stripped = StripSrgbClrTint(doc, w);
                    return pinned || stripped;
                },
                null,
                out object next)
                || next == null)
            {
                return;
            }

            shape = next;
            chart = TryGetChart(shape);
            if (chart != null && snap != null)
            {
                TryWriteAreaFill(chart, "ChartArea", snap.ChartAreaFillVisible, snap.ChartAreaFillRgb);
                TryWriteAreaFill(chart, "PlotArea", snap.PlotFillVisible, snap.PlotFillRgb);
                PourLog(warnings, "OOXML 钉皮：贴回后重钉绘图区/图表区填充");
                LogAxisFormats(chart, warnings, "钉皮后");
            }
        }

        private static bool SnapNeedsWppOoxmlChrome(ChartStyleSnap snap)
        {
            if (snap == null || snap.Series == null)
            {
                return false;
            }

            for (int i = 0; i < snap.Series.Count; i++)
            {
                SeriesStyleSnap one = snap.Series[i];
                if (one == null)
                {
                    continue;
                }

                if (SeriesFillWantsOoxml(one))
                {
                    return true;
                }

                if (one.Line != null && (one.Line.Rgb.HasValue || one.Line.Visible == false))
                {
                    return true;
                }

                if (one.MarkerForeRgb.HasValue || one.MarkerBackRgb.HasValue)
                {
                    return true;
                }

                if (one.Fill != null && one.Fill.SolidRgb.HasValue)
                {
                    return true;
                }

                if (one.PointFills == null)
                {
                    continue;
                }

                for (int p = 0; p < one.PointFills.Count; p++)
                {
                    if (one.PointFills[p] != null && one.PointFills[p].SolidRgb.HasValue)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool SeriesFillWantsOoxml(SeriesStyleSnap one)
        {
            if (one == null || one.Fill == null)
            {
                return false;
            }

            if (one.Fill.Stops != null && one.Fill.Stops.Count >= 2)
            {
                return one.Fill.FillType == MsoFillGradient
                    || one.Fill.Angle.HasValue
                    || one.Fill.Stops.Exists(s => s != null && s.Transparency > 0.005);
            }

            return false;
        }

        private static bool StripSrgbClrTint(XDocument chartDoc, List<string> warnings)
        {
            if (chartDoc == null)
            {
                return false;
            }

            int removed = 0;
            foreach (XElement srgb in chartDoc.Descendants(ANs + "srgbClr").ToList())
            {
                foreach (XElement child in srgb.Elements().ToList())
                {
                    string local = child.Name.LocalName;
                    if (string.Equals(local, "tint", StringComparison.Ordinal)
                        || string.Equals(local, "shade", StringComparison.Ordinal))
                    {
                        child.Remove();
                        removed++;
                    }
                }
            }

            if (removed == 0)
            {
                foreach (XElement srgb in chartDoc.Descendants().Where(e => e.Name.LocalName == "srgbClr").ToList())
                {
                    foreach (XElement child in srgb.Elements().ToList())
                    {
                        string local = child.Name.LocalName;
                        if (string.Equals(local, "tint", StringComparison.Ordinal)
                            || string.Equals(local, "shade", StringComparison.Ordinal))
                        {
                            child.Remove();
                            removed++;
                        }
                    }
                }
            }

            PourLog(warnings, "OOXML 剥tint：去掉 " + removed + " 个 tint/shade");
            return removed > 0;
        }

        private static readonly string[] SerChildOrder =
        {
            "idx", "order", "tx", "spPr", "invertIfNegative", "pictureOptions",
            "marker", "dPt", "dLbls", "trendline", "errBars",
            "cat", "val", "xVal", "yVal", "smooth", "shape", "bubbleSize", "extLst"
        };

        private static readonly string[] SpPrChildOrder =
        {
            "xfrm", "custGeom", "prstGeom",
            "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill",
            "ln", "effectLst", "effectDag", "scene3d", "sp3d", "extLst"
        };

        private static bool RewriteSeriesChromeFromSnap(XDocument chartDoc, ChartStyleSnap snap, List<string> warnings)
        {
            if (chartDoc == null || snap == null || snap.Series == null)
            {
                return false;
            }

            List<XElement> series = chartDoc.Descendants(CNs + "ser").ToList();
            if (series.Count == 0)
            {
                series = chartDoc.Descendants().Where(e => e.Name.LocalName == "ser").ToList();
            }

            int pinned = 0;
            int n = Math.Min(series.Count, snap.Series.Count);
            for (int i = 0; i < n; i++)
            {
                if (PinSerChrome(series[i], snap.Series[i]))
                {
                    pinned++;
                }
            }

            PourLog(warnings, "OOXML 钉皮：系列 " + pinned + "/" + n);
            return pinned > 0;
        }

        private static bool PinSerChrome(XElement ser, SeriesStyleSnap one)
        {
            if (ser == null || one == null)
            {
                return false;
            }

            bool changed = false;
            if (SeriesFillWantsOoxml(one))
            {
                XElement spPr = EnsureNamedChild(ser, CNs + "spPr", SerChildOrder);
                WriteGradFill(spPr, one.Fill);
                ClearPointFills(ser);
                changed = true;
            }
            else if (one.Fill != null
                && one.Fill.Visible == false
                && (one.Fill.Stops == null || one.Fill.Stops.Count < 2)
                && !IsLineLike(one.ChartType))
            {
                XElement spPr = EnsureNamedChild(ser, CNs + "spPr", SerChildOrder);
                RemoveFillKinds(spPr);
                EnsureNamedChild(spPr, ANs + "noFill", SpPrChildOrder);
                changed = true;
            }

            if (one.Line != null && (one.Line.Rgb.HasValue || one.Line.Visible == false || HasSaneWeight(one.Line.Weight)))
            {
                XElement spPr = EnsureNamedChild(ser, CNs + "spPr", SerChildOrder);
                WriteSeriesLine(spPr, one.Line);
                changed = true;
            }

            if (one.MarkerForeRgb.HasValue
                || one.MarkerBackRgb.HasValue
                || one.MarkerStyle.HasValue
                || one.MarkerSize.HasValue)
            {
                WriteMarker(ser, one);
                changed = true;
            }

            return changed;
        }

        private static void WriteGradFill(XElement spPr, FillSnap fill)
        {
            if (spPr == null || fill == null || fill.Stops == null || fill.Stops.Count < 2)
            {
                return;
            }

            RemoveFillKinds(spPr);
            var grad = new XElement(ANs + "gradFill");
            var gsLst = new XElement(ANs + "gsLst");
            for (int i = 0; i < fill.Stops.Count; i++)
            {
                GradientStopSnap stop = fill.Stops[i];
                if (stop == null)
                {
                    continue;
                }

                int pos = (int)Math.Round(Clamp01(stop.Position) * 100000);
                var gs = new XElement(ANs + "gs");
                gs.SetAttributeValue("pos", pos.ToString(CultureInfo.InvariantCulture));
                gs.Add(MakeSrgbClr(stop.Rgb, stop.Transparency));
                gsLst.Add(gs);
            }

            grad.Add(gsLst);
            if (fill.Angle.HasValue && IsSaneFillAngle(fill.Angle.Value))
            {
                int ang = (int)Math.Round(fill.Angle.Value * 60000);
                var lin = new XElement(ANs + "lin");
                lin.SetAttributeValue("ang", ang.ToString(CultureInfo.InvariantCulture));
                lin.SetAttributeValue("scaled", "1");
                grad.Add(lin);
            }

            InsertNamed(spPr, grad, SpPrChildOrder);
        }

        private static void WriteSeriesLine(XElement spPr, LineSnap line)
        {
            if (spPr == null || line == null)
            {
                return;
            }

            XElement ln = LocalChild(spPr, "ln");
            if (ln == null)
            {
                ln = new XElement(ANs + "ln");
                InsertNamed(spPr, ln, SpPrChildOrder);
            }

            if (line.Visible == false)
            {
                RemoveFillKinds(ln);
                EnsureNamedChild(ln, ANs + "noFill", new[] { "noFill", "solidFill", "gradFill", "prstDash", "extLst" });
                return;
            }

            if (HasSaneWeight(line.Weight))
            {
                int emu = (int)Math.Round(line.Weight.Value * 12700);
                if (emu < 1)
                {
                    emu = 12700;
                }

                ln.SetAttributeValue("w", emu.ToString(CultureInfo.InvariantCulture));
            }

            if (line.Rgb.HasValue)
            {
                RemoveFillKinds(ln);
                var solid = new XElement(ANs + "solidFill");
                solid.Add(MakeSrgbClr(line.Rgb.Value, 0));
                ln.AddFirst(solid);
            }
        }

        private static void WriteMarker(XElement ser, SeriesStyleSnap one)
        {
            XElement marker = EnsureNamedChild(ser, CNs + "marker", SerChildOrder);
            if (one.MarkerStyle.HasValue)
            {
                string symbol = MarkerFromXl(one.MarkerStyle.Value);
                XElement sym = EnsureNamedChild(marker, CNs + "symbol", new[] { "symbol", "size", "spPr", "extLst" });
                if (!string.IsNullOrEmpty(symbol))
                {
                    sym.SetAttributeValue("val", symbol);
                }
            }

            if (one.MarkerSize.HasValue && one.MarkerSize.Value > 0)
            {
                XElement size = EnsureNamedChild(marker, CNs + "size", new[] { "symbol", "size", "spPr", "extLst" });
                size.SetAttributeValue("val", one.MarkerSize.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (!one.MarkerForeRgb.HasValue && !one.MarkerBackRgb.HasValue)
            {
                return;
            }

            XElement spPr = EnsureNamedChild(marker, CNs + "spPr", new[] { "symbol", "size", "spPr", "extLst" });
            if (one.MarkerBackRgb.HasValue)
            {
                RemoveFillKinds(spPr);
                var solid = new XElement(ANs + "solidFill");
                solid.Add(MakeSrgbClr(one.MarkerBackRgb.Value, 0));
                InsertNamed(spPr, solid, SpPrChildOrder);
            }

            if (one.MarkerForeRgb.HasValue)
            {
                XElement ln = LocalChild(spPr, "ln");
                if (ln == null)
                {
                    ln = new XElement(ANs + "ln");
                    InsertNamed(spPr, ln, SpPrChildOrder);
                }

                RemoveFillKinds(ln);
                var solid = new XElement(ANs + "solidFill");
                solid.Add(MakeSrgbClr(one.MarkerForeRgb.Value, 0));
                ln.AddFirst(solid);
            }
        }

        private static void ClearPointFills(XElement ser)
        {
            foreach (XElement dpt in ser.Elements().Where(e => e.Name.LocalName == "dPt").ToList())
            {
                XElement spPr = LocalChild(dpt, "spPr");
                if (spPr != null)
                {
                    RemoveFillKinds(spPr);
                }
            }
        }

        private static void TryEnrichSnapFromOoxml(object shape, ChartStyleSnap snap, List<string> warnings = null)
        {
            if (shape == null || snap == null || snap.Series == null || snap.Series.Count == 0)
            {
                return;
            }

            if (!TryReadLiveChartXml(shape, warnings, out XDocument chartDoc) || chartDoc == null)
            {
                return;
            }

            List<XElement> series = chartDoc.Descendants(CNs + "ser").ToList();
            if (series.Count == 0)
            {
                series = chartDoc.Descendants().Where(e => e.Name.LocalName == "ser").ToList();
            }

            int filled = 0;
            int n = Math.Min(series.Count, snap.Series.Count);
            for (int i = 0; i < n; i++)
            {
                if (MergeSerChromeFromXml(snap.Series[i], series[i]))
                {
                    filled++;
                }
            }

            PourLog(warnings, "OOXML 读皮：补系列 " + filled + "/" + n);
        }

        private static bool MergeSerChromeFromXml(SeriesStyleSnap one, XElement ser)
        {
            if (one == null || ser == null)
            {
                return false;
            }

            bool changed = false;
            XElement spPr = LocalChild(ser, "spPr");
            FillSnap xmlFill = ReadXmlFill(spPr);
            if (xmlFill != null && xmlFill.Stops != null && xmlFill.Stops.Count >= 2)
            {
                if (one.Fill == null)
                {
                    one.Fill = xmlFill;
                }
                else
                {
                    one.Fill.Stops = xmlFill.Stops;
                    one.Fill.FillType = MsoFillGradient;
                    one.Fill.Visible = true;
                    if (xmlFill.Angle.HasValue)
                    {
                        one.Fill.Angle = xmlFill.Angle;
                    }
                }

                changed = true;
            }
            else if (one.Fill != null
                && one.Fill.Angle.HasValue
                && !IsSaneFillAngle(one.Fill.Angle.Value))
            {
                one.Fill.Angle = null;
                changed = true;
            }

            LineSnap xmlLine = ReadXmlLine(spPr);
            bool wantLine = IsLineLike(one.ChartType) || one.Line != null;
            if (xmlLine != null && wantLine)
            {
                if (one.Line == null)
                {
                    one.Line = xmlLine;
                    changed = true;
                }
                else
                {
                    if (!one.Line.Rgb.HasValue && xmlLine.Rgb.HasValue)
                    {
                        one.Line.Rgb = xmlLine.Rgb;
                        one.Line.Visible = true;
                        changed = true;
                    }

                    if (!HasSaneWeight(one.Line.Weight) && HasSaneWeight(xmlLine.Weight))
                    {
                        one.Line.Weight = xmlLine.Weight;
                        changed = true;
                    }
                }
            }

            XElement marker = LocalChild(ser, "marker");
            bool wantMarker = one.MarkerStyle.HasValue
                || one.MarkerSize.HasValue
                || one.MarkerForeRgb.HasValue
                || one.MarkerBackRgb.HasValue
                || IsLineLike(one.ChartType);
            if (marker != null && wantMarker)
            {
                XElement mSp = LocalChild(marker, "spPr");
                if (!one.MarkerBackRgb.HasValue)
                {
                    int? back = ReadSolidRgb(mSp);
                    if (back.HasValue)
                    {
                        one.MarkerBackRgb = back;
                        changed = true;
                    }
                }

                if (!one.MarkerForeRgb.HasValue)
                {
                    int? fore = ReadSolidRgb(LocalChild(mSp, "ln"));
                    if (fore.HasValue)
                    {
                        one.MarkerForeRgb = fore;
                        changed = true;
                    }
                }

                if (!one.MarkerStyle.HasValue)
                {
                    string symbol = (string)(LocalChild(marker, "symbol") == null
                        ? null
                        : LocalChild(marker, "symbol").Attribute("val"));
                    if (!string.IsNullOrWhiteSpace(symbol))
                    {
                        one.MarkerStyle = MarkerToXl(symbol);
                        changed = true;
                    }
                }

                if (!one.MarkerSize.HasValue)
                {
                    string sizeRaw = (string)(LocalChild(marker, "size") == null
                        ? null
                        : LocalChild(marker, "size").Attribute("val"));
                    if (int.TryParse(sizeRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int ms)
                        && ms > 0)
                    {
                        one.MarkerSize = ms;
                        changed = true;
                    }
                }
            }

            return changed;
        }

        private static FillSnap ReadXmlFill(XElement spPr)
        {
            XElement grad = LocalChild(spPr, "gradFill");
            if (grad == null)
            {
                return null;
            }

            XElement gsLst = LocalChild(grad, "gsLst");
            if (gsLst == null)
            {
                return null;
            }

            var stops = new List<GradientStopSnap>();
            foreach (XElement gs in gsLst.Elements().Where(e => e.Name.LocalName == "gs"))
            {
                if (!TryParseGradPos((string)gs.Attribute("pos"), out double pos))
                {
                    continue;
                }

                if (!TryReadSrgb(gs, out int rgb, out double trans))
                {
                    continue;
                }

                stops.Add(new GradientStopSnap
                {
                    Position = pos,
                    Rgb = rgb,
                    Transparency = trans
                });
            }

            if (stops.Count < 2)
            {
                return null;
            }

            var fill = new FillSnap
            {
                Visible = true,
                FillType = MsoFillGradient,
                Stops = stops
            };
            XElement lin = LocalChild(grad, "lin");
            if (lin != null
                && int.TryParse((string)lin.Attribute("ang"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int ang))
            {
                double deg = ang / 60000.0;
                if (IsSaneFillAngle(deg))
                {
                    fill.Angle = deg;
                }
            }

            return fill;
        }

        private static LineSnap ReadXmlLine(XElement spPr)
        {
            XElement ln = LocalChild(spPr, "ln");
            if (ln == null)
            {
                return null;
            }

            var line = new LineSnap();
            if (LocalChild(ln, "noFill") != null)
            {
                line.Visible = false;
                return line;
            }

            if (TryReadSrgb(ln, out int rgb, out _))
            {
                line.Rgb = rgb;
                line.Visible = true;
            }

            if (int.TryParse((string)ln.Attribute("w"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int emu)
                && emu > 0)
            {
                double pt = emu / 12700.0;
                if (HasSaneWeight(pt))
                {
                    line.Weight = pt;
                }
            }

            return line.Rgb.HasValue || line.Weight.HasValue || line.Visible.HasValue ? line : null;
        }

        private static int? ReadSolidRgb(XElement parent)
        {
            if (parent == null)
            {
                return null;
            }

            XElement solid = LocalChild(parent, "solidFill");
            if (solid != null && TryReadSrgb(solid, out int rgb, out _))
            {
                return rgb;
            }

            if (TryReadSrgb(parent, out int direct, out _))
            {
                return direct;
            }

            return null;
        }

        private static bool TryReadSrgb(XElement parent, out int rgb, out double transparency)
        {
            rgb = 0;
            transparency = 0;
            XElement srgb = LocalChild(parent, "srgbClr")
                ?? parent.Elements().FirstOrDefault(e => e.Name.LocalName == "srgbClr");
            if (srgb == null)
            {
                XElement solid = LocalChild(parent, "solidFill");
                srgb = solid == null ? null : LocalChild(solid, "srgbClr");
            }

            if (srgb == null)
            {
                return false;
            }

            string val = (string)srgb.Attribute("val");
            if (string.IsNullOrWhiteSpace(val) || !TryParseHexToOffice("#" + val.Trim(), out rgb))
            {
                return false;
            }

            XElement alpha = LocalChild(srgb, "alpha");
            if (alpha != null
                && int.TryParse((string)alpha.Attribute("val"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int a))
            {
                transparency = 1.0 - (a / 100000.0);
                if (transparency < 0)
                {
                    transparency = 0;
                }

                if (transparency > 1)
                {
                    transparency = 1;
                }
            }

            return true;
        }

        private static bool TryParseGradPos(string raw, out double pos01)
        {
            pos01 = 0;
            if (string.IsNullOrWhiteSpace(raw))
            {
                return false;
            }

            string t = raw.Trim();
            if (t.EndsWith("%", StringComparison.Ordinal))
            {
                if (!double.TryParse(t.Substring(0, t.Length - 1), NumberStyles.Float, CultureInfo.InvariantCulture, out double pct))
                {
                    return false;
                }

                pos01 = Clamp01(pct / 100.0);
                return true;
            }

            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out double v))
            {
                return false;
            }

            pos01 = v > 1.0001 ? Clamp01(v / 100000.0) : Clamp01(v);
            return true;
        }

        private static XElement MakeSrgbClr(int officeRgb, double transparency)
        {
            string hex = OfficeRgbToHex(officeRgb);
            if (hex != null && hex.StartsWith("#", StringComparison.Ordinal))
            {
                hex = hex.Substring(1);
            }

            var el = new XElement(ANs + "srgbClr");
            el.SetAttributeValue("val", hex);
            if (transparency > 0.005)
            {
                int alpha = (int)Math.Round((1.0 - transparency) * 100000);
                if (alpha < 0)
                {
                    alpha = 0;
                }

                if (alpha > 100000)
                {
                    alpha = 100000;
                }

                el.Add(new XElement(ANs + "alpha", new XAttribute("val", alpha.ToString(CultureInfo.InvariantCulture))));
            }

            return el;
        }

        private static void RemoveFillKinds(XElement parent)
        {
            if (parent == null)
            {
                return;
            }

            foreach (XElement child in parent.Elements().ToList())
            {
                string local = child.Name.LocalName;
                if (local == "solidFill" || local == "gradFill" || local == "noFill"
                    || local == "blipFill" || local == "pattFill" || local == "grpFill")
                {
                    child.Remove();
                }
            }
        }

        private static XElement EnsureNamedChild(XElement parent, XName name, string[] order)
        {
            XElement existing = LocalChild(parent, name.LocalName);
            if (existing != null)
            {
                return existing;
            }

            var created = new XElement(name);
            InsertNamed(parent, created, order);
            return created;
        }

        private static void InsertNamed(XElement parent, XElement created, string[] order)
        {
            if (parent == null || created == null)
            {
                return;
            }

            int want = order == null ? -1 : Array.IndexOf(order, created.Name.LocalName);
            if (want >= 0)
            {
                foreach (XElement child in parent.Elements())
                {
                    int at = Array.IndexOf(order, child.Name.LocalName);
                    if (at > want)
                    {
                        child.AddBeforeSelf(created);
                        return;
                    }
                }
            }

            parent.Add(created);
        }

        private static XElement LocalChild(XElement parent, string local)
        {
            if (parent == null || string.IsNullOrEmpty(local))
            {
                return null;
            }

            return parent.Elements().FirstOrDefault(e => e.Name.LocalName == local);
        }

        private static bool IsSaneFillAngle(double deg)
        {
            return deg >= -360 && deg <= 360;
        }

        private static bool HasSaneWeight(double? weight)
        {
            return weight.HasValue && !IsPhantomWeight(weight.Value);
        }

        private static bool TryReadLiveChartXml(object shape, List<string> warnings, out XDocument chartDoc)
        {
            chartDoc = null;
            if (shape == null)
            {
                return false;
            }

            object slide = TryGetShapeSlide(shape);
            object pres = TryGetShapePresentation(shape);
            if (slide == null || pres == null)
            {
                return false;
            }

            object app = WppCom.GetProperty(pres, "Application");
            if (!LooksLikeWppApp(app))
            {
                return false;
            }

            int shapeId = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
            int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
            if (shapeId < 1 || slideIndex < 1)
            {
                return false;
            }

            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "ew-wpp-chart-read-" + Guid.NewGuid().ToString("N") + ".pptx");
            try
            {
                try
                {
                    WppCom.Invoke(pres, "SaveCopyAs", tempPath);
                }
                catch (Exception)
                {
                    WppCom.Invoke(pres, "SaveCopyAs", tempPath, 24);
                }

                if (!File.Exists(tempPath))
                {
                    return false;
                }

                return TryReadChartDocFromPptx(tempPath, slideIndex, shapeId, warnings, "读皮", out chartDoc);
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML 读皮失败: " + FormatComError(ex));
                return false;
            }
            finally
            {
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool TryReadChartDocFromPptx(
            string pptxPath,
            int slideIndex,
            int shapeId,
            List<string> warnings,
            string tag,
            out XDocument chartDoc)
        {
            chartDoc = null;
            try
            {
                using (ZipArchive zip = ZipFile.OpenRead(pptxPath))
                {
                    return TryOpenChartDocInZip(zip, slideIndex, shapeId, warnings, tag, out chartDoc, out _);
                }
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML " + tag + "打开包失败: " + FormatComError(ex));
                return false;
            }
        }

        private static bool TryEditChartXmlAndCopyBack(
            object shape,
            List<string> warnings,
            string tag,
            Func<XDocument, List<string>, bool> rewrite,
            Func<object, bool> acceptCopy,
            out object newShape)
        {
            newShape = null;
            if (shape == null || rewrite == null)
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
                PourLog(warnings, "OOXML " + tag + "：进入");
                slide = TryGetShapeSlide(shape);
                pres = TryGetShapePresentation(shape);
                if (slide == null || pres == null)
                {
                    PourLog(warnings, "OOXML " + tag + "：取不到 slide/presentation");
                    return false;
                }

                app = WppCom.GetProperty(pres, "Application");
                string appName = TryPropString(app, "Name") ?? "";
                PourLog(warnings, "OOXML " + tag + "：宿主 Name=" + appName);
                if (!LooksLikeWppApp(app))
                {
                    PourLog(warnings, "OOXML " + tag + "：非 WPP 宿主（" + appName + "），跳过");
                    return false;
                }

                int shapeId = Convert.ToInt32(WppCom.GetProperty(shape, "Id") ?? 0);
                int slideIndex = Convert.ToInt32(WppCom.GetProperty(slide, "SlideIndex") ?? 0);
                PourLog(warnings, "OOXML " + tag + "：shapeId=" + shapeId + " slideIndex=" + slideIndex);
                if (shapeId < 1 || slideIndex < 1)
                {
                    PourLog(warnings, "OOXML " + tag + "：shapeId/slideIndex 非法 " + shapeId + "/" + slideIndex);
                    return false;
                }

                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "ew-wpp-chart-" + Guid.NewGuid().ToString("N") + ".pptx");
                PourLog(warnings, "OOXML " + tag + "：SaveCopyAs → " + tempPath);
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
                    PourLog(warnings, "OOXML " + tag + "：SaveCopyAs 未落盘 " + tempPath);
                    return false;
                }

                PourLog(warnings, "OOXML " + tag + "：副本已落盘 " + tempPath + " size=" + new FileInfo(tempPath).Length);

                if (!TryRewriteChartXmlInPackage(tempPath, slideIndex, shapeId, rewrite, warnings, tag))
                {
                    return false;
                }

                PourLog(warnings, "OOXML " + tag + "：打开副本 " + tempPath);
                object presentations = WppCom.GetProperty(app, "Presentations");
                copyPres = TryOpenCopyPresentation(presentations, tempPath, warnings);
                if (copyPres == null)
                {
                    PourLog(warnings, "OOXML " + tag + "：打开副本失败");
                    return false;
                }

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
                    PourLog(warnings, "OOXML " + tag + "：副本里找不到 shapeId=" + shapeId);
                    return false;
                }

                if (acceptCopy != null && !acceptCopy(copyChartShape))
                {
                    PourLog(warnings, "OOXML " + tag + "：副本图验收未过");
                    return false;
                }

                PourLog(warnings, "OOXML " + tag + "：Copy 副本图");
                WppCom.Invoke(copyChartShape, "Copy");
                object shapes = WppCom.GetProperty(slide, "Shapes");
                object pasted = TryPasteChart(shapes, warnings);
                if (pasted == null)
                {
                    PourLog(warnings, "OOXML " + tag + "：粘贴回原页失败");
                    return false;
                }

                TryCopyBox(shape, pasted);
                if (acceptCopy != null && !acceptCopy(pasted))
                {
                    PourLog(warnings, "OOXML " + tag + "：贴回后验收未过");
                    TryDelete(pasted);
                    return false;
                }

                TryDelete(shape);
                newShape = pasted;
                PourLog(warnings, "OOXML " + tag + "完成");
                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML " + tag + "失败: " + FormatComError(ex));
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
            Func<XDocument, List<string>, bool> rewrite,
            List<string> warnings,
            string tag)
        {
            string tempOut = pptxPath + ".fix";
            try
            {
                using (ZipArchive zip = ZipFile.Open(pptxPath, ZipArchiveMode.Update))
                {
                    if (!TryOpenChartDocInZip(
                        zip,
                        slideIndex,
                        shapeId,
                        warnings,
                        tag,
                        out XDocument chartDoc,
                        out string chartPart))
                    {
                        return false;
                    }

                    if (!rewrite(chartDoc, warnings))
                    {
                        return false;
                    }

                    ZipArchiveEntry chartEntry = FindEntry(zip, chartPart);
                    if (chartEntry == null)
                    {
                        return false;
                    }

                    string fullName = chartEntry.FullName;
                    chartEntry.Delete();
                    ZipArchiveEntry fresh = zip.CreateEntry(fullName);
                    using (Stream s = fresh.Open())
                    {
                        chartDoc.Save(s);
                    }

                    PourLog(warnings, "OOXML " + tag + "：已改 " + chartPart);
                }

                return true;
            }
            catch (Exception ex)
            {
                PourLog(warnings, "OOXML " + tag + "改包失败: " + FormatComError(ex));
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

        private static bool TryOpenChartDocInZip(
            ZipArchive zip,
            int slideIndex,
            int shapeId,
            List<string> warnings,
            string tag,
            out XDocument chartDoc,
            out string chartPart)
        {
            chartDoc = null;
            chartPart = null;
            string slidePart = "ppt/slides/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml";
            string slideRels = "ppt/slides/_rels/slide" + slideIndex.ToString(CultureInfo.InvariantCulture) + ".xml.rels";
            ZipArchiveEntry slideEntry = FindEntry(zip, slidePart);
            ZipArchiveEntry relsEntry = FindEntry(zip, slideRels);
            if (slideEntry == null || relsEntry == null)
            {
                PourLog(warnings, "OOXML " + tag + "：缺 slide 部件 " + slidePart);
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
                PourLog(warnings, "OOXML " + tag + "：slide 里找不到 shapeId=" + shapeId + " 的 chart 关系");
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
                PourLog(warnings, "OOXML " + tag + "：rels 里找不到 " + relId);
                return false;
            }

            chartPart = NormalizeChartPartPath(target);
            ZipArchiveEntry chartEntry = FindEntry(zip, chartPart);
            if (chartEntry == null)
            {
                PourLog(warnings, "OOXML " + tag + "：缺 chart 部件 " + chartPart);
                return false;
            }

            using (Stream s = chartEntry.Open())
            {
                chartDoc = XDocument.Load(s);
            }

            return chartDoc != null;
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
