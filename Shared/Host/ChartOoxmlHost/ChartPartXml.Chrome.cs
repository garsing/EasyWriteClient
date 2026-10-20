using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.ChartOoxmlHost
{
    internal static partial class ChartPartXml
    {
        private static void WriteChartChrome(XDocument chartDoc, SeriesGrid grid)
        {
            if (chartDoc == null || grid == null)
            {
                return;
            }

            WriteLegend(chartDoc, grid.Legend, grid.LegendFontColor);
            XElement plot = chartDoc.Descendants(C + "plotArea").FirstOrDefault();
            WriteAreaFill(plot, grid.PlotColor);
            WriteAreaFill(chartDoc.Root, grid.ChartAreaColor);
            XElement primary = FindValAx(chartDoc, false);
            if (primary != null)
            {
                string yTitle = !string.IsNullOrEmpty(grid.AxisYTitle)
                    ? grid.AxisYTitle
                    : grid.AxisY != null ? grid.AxisY.Title : null;
                WriteAxisTitle(primary, yTitle);
                WriteScale(primary, grid.YMin, grid.YMax, grid.YMajor);
                WriteAxisStyle(primary, grid.AxisY);
                WriteGridlines(primary, grid.Gridlines);
            }

            XElement cat = plot != null ? plot.Elements(C + "catAx").FirstOrDefault() : null;
            string xTitle = !string.IsNullOrEmpty(grid.AxisXTitle)
                ? grid.AxisXTitle
                : grid.AxisX != null ? grid.AxisX.Title : null;
            WriteAxisTitle(cat, xTitle);
            WriteAxisStyle(cat, StyleFromAxisX(grid.AxisX));
            WriteChartStyle(chartDoc, grid.ChartStyle);
            WriteBarGap(chartDoc, grid.GapWidth, grid.Overlap);

            XElement secondary = FindValAx(chartDoc, true);
            if (secondary != null)
            {
                string y2Title = !string.IsNullOrEmpty(grid.AxisY2Title)
                    ? grid.AxisY2Title
                    : grid.AxisY2 != null ? grid.AxisY2.Title : null;
                WriteAxisTitle(secondary, y2Title);
                WriteAxisStyle(secondary, grid.AxisY2);
            }
        }

        private static void WriteSeriesChrome(XElement ser, SeriesCol col)
        {
            if (ser == null || col == null)
            {
                return;
            }

            if (IsNone(col.Color))
            {
                WriteSeriesNoFill(ser);
            }
            else if (!WriteGradient(ser, col.FillGradient, col.FillAngle))
            {
                string fill = NormColor(col.Color);
                if (!string.IsNullOrEmpty(fill))
                {
                    PinSerColor(ser, fill);
                }
            }

            string line = NormColor(col.Line);
            if (!string.IsNullOrEmpty(line) || !string.IsNullOrEmpty(col.LineWeight))
            {
                WriteLn(ser, line, col.LineWeight);
            }

            WriteMarker(ser, col);

            if (col.ShowDataLabels.HasValue
                || col.ShowValue.HasValue
                || col.ShowPercentage.HasValue
                || !string.IsNullOrEmpty(col.LabelPosition)
                || !string.IsNullOrEmpty(col.LabelFont)
                || !string.IsNullOrEmpty(col.LabelSize)
                || !string.IsNullOrEmpty(col.LabelColor)
                || !string.IsNullOrEmpty(col.LabelFormat))
            {
                XElement labels = ser.Element(C + "dLbls");
                if (labels == null)
                {
                    labels = new XElement(C + "dLbls");
                    XElement ext = ser.Element(C + "extLst");
                    if (ext != null)
                    {
                        ext.AddBeforeSelf(labels);
                    }
                    else
                    {
                        ser.Add(labels);
                    }
                }

                string labelPos = col.LabelPosition;
                string parentName = ser.Parent == null ? "" : ser.Parent.Name.LocalName;
                if ((parentName == "barChart" || parentName == "bar3DChart")
                    && (string.Equals(labelPos, "above", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(labelPos, "below", StringComparison.OrdinalIgnoreCase)))
                {
                    labelPos = "outside";
                }

                string pos = LabelPosToOoxml(labelPos);
                if (!string.IsNullOrEmpty(pos))
                {
                    SetVal(labels, "dLblPos", pos);
                }

                if (col.ShowDataLabels == false
                    && !col.ShowValue.HasValue
                    && !col.ShowPercentage.HasValue)
                {
                    SetVal(labels, "delete", "1");
                }
                else
                {
                    labels.Element(C + "delete")?.Remove();
                    bool? showVal = col.ShowValue;
                    if (!showVal.HasValue && col.ShowDataLabels == true)
                    {
                        showVal = true;
                    }

                    if (showVal.HasValue)
                    {
                        SetVal(labels, "showVal", showVal.Value ? "1" : "0");
                        SetVal(labels, "showCatName", "0");
                        SetVal(labels, "showSerName", "0");
                    }

                    if (col.ShowPercentage.HasValue)
                    {
                        SetVal(labels, "showPercent", col.ShowPercentage.Value ? "1" : "0");
                    }
                    else if (col.ShowDataLabels == true && !col.ShowValue.HasValue)
                    {
                        SetVal(labels, "showPercent", "0");
                    }
                }

                WriteLabelNumFmt(labels, col.LabelFormat);
                WriteLabelText(labels, col.LabelFont, col.LabelSize, col.LabelColor);
            }
        }

        private static void WriteLabelNumFmt(XElement labels, string format)
        {
            if (labels == null || string.IsNullOrEmpty(format))
            {
                return;
            }

            labels.Element(C + "numFmt")?.Remove();
            var num = new XElement(C + "numFmt");
            num.SetAttributeValue("formatCode", format);
            num.SetAttributeValue("sourceLinked", "0");
            labels.AddFirst(num);
        }

        private static void WriteLabelText(XElement labels, string font, string size, string color)
        {
            if (labels == null
                || (string.IsNullOrEmpty(font) && string.IsNullOrEmpty(size) && string.IsNullOrEmpty(color)))
            {
                return;
            }

            labels.Element(C + "txPr")?.Remove();
            var def = new XElement(A + "defRPr");
            double pt;
            if (!string.IsNullOrEmpty(size)
                && double.TryParse(size, NumberStyles.Float, CultureInfo.InvariantCulture, out pt))
            {
                int hundredths = (int)Math.Round(pt * 100.0);
                def.SetAttributeValue("sz", hundredths.ToString(CultureInfo.InvariantCulture));
            }

            string hex = NormColor(color);
            if (!string.IsNullOrEmpty(hex))
            {
                def.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", hex))));
            }

            if (!string.IsNullOrEmpty(font))
            {
                def.Add(new XElement(A + "latin", new XAttribute("typeface", font)));
                def.Add(new XElement(A + "ea", new XAttribute("typeface", font)));
            }

            var tx = new XElement(
                C + "txPr",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p", new XElement(A + "pPr", def), new XElement(A + "endParaRPr")));
            XElement num = labels.Element(C + "numFmt");
            if (num != null)
            {
                num.AddAfterSelf(tx);
            }
            else
            {
                labels.AddFirst(tx);
            }
        }

        private static bool WriteGradient(XElement ser, string gradient, string angle)
        {
            if (ser == null || string.IsNullOrEmpty(gradient))
            {
                return false;
            }

            string[] parts = gradient.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            var stops = new List<XElement>();
            for (int i = 0; i < parts.Length; i++)
            {
                XElement gs = GradientStop(parts[i]);
                if (gs != null)
                {
                    stops.Add(gs);
                }
            }

            if (stops.Count < 2)
            {
                return false;
            }

            XElement spPr = ser.Element(C + "spPr");
            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                ser.Add(spPr);
            }

            spPr.Element(A + "solidFill")?.Remove();
            spPr.Element(A + "gradFill")?.Remove();
            spPr.Element(A + "noFill")?.Remove();
            var grad = new XElement(A + "gradFill", new XElement(A + "gsLst", stops));
            double deg;
            if (!string.IsNullOrEmpty(angle)
                && double.TryParse(angle, NumberStyles.Float, CultureInfo.InvariantCulture, out deg))
            {
                int ang = (int)Math.Round(deg * 60000.0);
                var lin = new XElement(A + "lin");
                lin.SetAttributeValue("ang", ang.ToString(CultureInfo.InvariantCulture));
                lin.SetAttributeValue("scaled", "1");
                grad.Add(lin);
            }

            spPr.AddFirst(grad);
            return true;
        }

        private static XElement GradientStop(string part)
        {
            if (string.IsNullOrEmpty(part))
            {
                return null;
            }

            int colon = part.IndexOf(':');
            if (colon <= 0)
            {
                return null;
            }

            double pos;
            if (!double.TryParse(part.Substring(0, colon), NumberStyles.Float, CultureInfo.InvariantCulture, out pos))
            {
                return null;
            }

            string rest = part.Substring(colon + 1);
            int at = rest.LastIndexOf('@');
            string color = at >= 0 ? rest.Substring(0, at) : rest;
            double transparency = 0;
            if (at >= 0)
            {
                double.TryParse(rest.Substring(at + 1), NumberStyles.Float, CultureInfo.InvariantCulture, out transparency);
            }

            string hex = NormColor(color);
            if (string.IsNullOrEmpty(hex))
            {
                return null;
            }

            int gsPos = (int)Math.Round(Math.Max(0, Math.Min(1, pos)) * 100000.0);
            int alpha = (int)Math.Round((1.0 - transparency) * 100000.0);
            if (alpha < 0)
            {
                alpha = 0;
            }

            if (alpha > 100000)
            {
                alpha = 100000;
            }

            var srgb = new XElement(A + "srgbClr", new XAttribute("val", hex));
            srgb.Add(new XElement(A + "alpha", new XAttribute("val", alpha.ToString(CultureInfo.InvariantCulture))));
            var gs = new XElement(A + "gs", new XAttribute("pos", gsPos.ToString(CultureInfo.InvariantCulture)));
            gs.Add(srgb);
            return gs;
        }

        private static void WriteAreaFill(XElement owner, string color)
        {
            if (owner == null || string.IsNullOrEmpty(color))
            {
                return;
            }

            XElement spPr = owner.Element(C + "spPr");
            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                owner.AddFirst(spPr);
            }

            foreach (XElement old in spPr.Elements(A + "solidFill").ToList())
            {
                old.Remove();
            }

            foreach (XElement old in spPr.Elements(A + "noFill").ToList())
            {
                old.Remove();
            }

            if (IsNone(color))
            {
                spPr.AddFirst(new XElement(A + "noFill"));
                return;
            }

            string hex = NormColor(color);
            if (string.IsNullOrEmpty(hex))
            {
                return;
            }

            spPr.AddFirst(new XElement(
                A + "solidFill",
                new XElement(A + "srgbClr", new XAttribute("val", hex))));
        }

        private static string ReadAreaFill(XElement owner)
        {
            XElement spPr = owner?.Element(C + "spPr");
            if (spPr == null)
            {
                return null;
            }

            if (spPr.Element(A + "noFill") != null)
            {
                return "none";
            }

            string hex = spPr.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
            return HashColor(hex);
        }

        private static void WriteLegend(XDocument chartDoc, string legend, string fontColor)
        {
            if (string.IsNullOrEmpty(legend) && string.IsNullOrEmpty(NormColor(fontColor)))
            {
                return;
            }

            XElement chart = chartDoc.Descendants(C + "chart").FirstOrDefault();
            if (chart == null)
            {
                return;
            }

            XElement node = chart.Elements(C + "legend").FirstOrDefault();
            if (IsNone(legend))
            {
                node?.Remove();
                return;
            }

            string pos = LegendPos(legend);
            if (node == null)
            {
                if (string.IsNullOrEmpty(pos) && string.IsNullOrEmpty(NormColor(fontColor)))
                {
                    return;
                }

                node = new XElement(C + "legend");
                XElement plot = chart.Element(C + "plotArea");
                if (plot != null)
                {
                    plot.AddAfterSelf(node);
                }
                else
                {
                    chart.Add(node);
                }
            }

            if (!string.IsNullOrEmpty(pos))
            {
                SetVal(node, "legendPos", pos);
                SetVal(node, "overlay", "0");
            }

            WriteLegendFont(node, fontColor);
        }

        private static void WriteLegendFont(XElement legend, string color)
        {
            string hex = NormColor(color);
            if (legend == null || string.IsNullOrEmpty(hex))
            {
                return;
            }

            legend.Element(C + "txPr")?.Remove();
            var def = new XElement(
                A + "defRPr",
                new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", hex))));
            var txPr = new XElement(
                C + "txPr",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p", new XElement(A + "pPr", def), new XElement(A + "endParaRPr")));
            XElement ext = legend.Element(C + "extLst");
            if (ext != null)
            {
                ext.AddBeforeSelf(txPr);
            }
            else
            {
                legend.Add(txPr);
            }
        }

        private static string ReadLegendFontColor(XDocument chartDoc)
        {
            XElement legend = chartDoc?.Descendants(C + "chart").FirstOrDefault()?.Elements(C + "legend").FirstOrDefault();
            return legend?.Element(C + "txPr")?.Descendants(A + "defRPr").FirstOrDefault()
                ?.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        }

        private static void WriteAxisTitle(XElement axis, string title)
        {
            if (axis == null || string.IsNullOrEmpty(title))
            {
                return;
            }

            foreach (XElement old in axis.Elements(C + "title").ToList())
            {
                old.Remove();
            }

            axis.Add(new XElement(
                C + "title",
                new XElement(
                    C + "tx",
                    new XElement(
                        C + "rich",
                        new XElement(A + "bodyPr"),
                        new XElement(A + "lstStyle"),
                        new XElement(
                            A + "p",
                            new XElement(
                                A + "r",
                                new XElement(A + "rPr", new XAttribute("lang", "zh-CN")),
                                new XElement(A + "t", title))))),
                new XElement(C + "overlay", new XAttribute("val", "0"))));
        }

        private static void WriteScale(XElement axis, double? min, double? max, double? major)
        {
            if (axis == null || (!min.HasValue && !max.HasValue && !major.HasValue))
            {
                return;
            }

            XElement scaling = axis.Element(C + "scaling");
            if (scaling == null)
            {
                scaling = new XElement(
                    C + "scaling",
                    new XElement(C + "orientation", new XAttribute("val", "minMax")));
                axis.AddFirst(scaling);
            }

            if (max.HasValue)
            {
                SetVal(scaling, "max", max.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (min.HasValue)
            {
                SetVal(scaling, "min", min.Value.ToString(CultureInfo.InvariantCulture));
            }

            XElement maxEl = scaling.Element(C + "max");
            XElement minEl = scaling.Element(C + "min");
            if (maxEl != null && minEl != null && minEl.IsBefore(maxEl))
            {
                maxEl.Remove();
                minEl.AddBeforeSelf(maxEl);
            }

            if (major.HasValue)
            {
                SetVal(axis, "majorUnit", major.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static void WriteAxisStyle(XElement axis, AxisYInfo style)
        {
            if (axis == null || style == null)
            {
                return;
            }

            string pos = TickPosToOoxml(style.TickPosition);
            if (string.IsNullOrEmpty(pos))
            {
                if (IsFalse(style.Visible))
                {
                    pos = "none";
                }
                else if (IsTrue(style.Visible))
                {
                    pos = "nextTo";
                }
            }

            if (!string.IsNullOrEmpty(pos))
            {
                SetVal(axis, "tickLblPos", pos);
            }

            string major = TickMarkToOoxml(style.MajorTick);
            if (!string.IsNullOrEmpty(major))
            {
                SetVal(axis, "majorTickMark", major);
            }

            string minor = TickMarkToOoxml(style.MinorTick);
            if (!string.IsNullOrEmpty(minor))
            {
                SetVal(axis, "minorTickMark", minor);
            }

            if (!string.IsNullOrEmpty(style.Format))
            {
                XElement num = axis.Element(C + "numFmt");
                if (num == null)
                {
                    num = new XElement(C + "numFmt");
                    axis.Add(num);
                }

                num.SetAttributeValue("formatCode", style.Format);
                num.SetAttributeValue("sourceLinked", "0");
            }

            if (IsTrue(style.Grid) || !string.IsNullOrEmpty(NormColor(style.GridColor)))
            {
                XElement grid = axis.Element(C + "majorGridlines");
                if (grid == null)
                {
                    grid = new XElement(C + "majorGridlines");
                    axis.Add(grid);
                }

                string gridColor = NormColor(style.GridColor);
                if (!string.IsNullOrEmpty(gridColor))
                {
                    WriteLn(grid, gridColor, null);
                }
            }
            else if (IsFalse(style.Grid))
            {
                axis.Element(C + "majorGridlines")?.Remove();
            }

            string line = NormColor(style.Line);
            if (!string.IsNullOrEmpty(line) || !string.IsNullOrEmpty(style.LineWeight))
            {
                WriteLn(axis, line, style.LineWeight);
            }

            if (!string.IsNullOrEmpty(style.TickFont)
                || !string.IsNullOrEmpty(style.TickSize)
                || !string.IsNullOrEmpty(style.TickColor))
            {
                WriteTickText(axis, style);
            }
        }

        private static void WriteTickText(XElement axis, AxisYInfo style)
        {
            foreach (XElement old in axis.Elements(C + "txPr").ToList())
            {
                old.Remove();
            }

            var def = new XElement(A + "defRPr");
            double size;
            if (!string.IsNullOrEmpty(style.TickSize)
                && double.TryParse(style.TickSize, NumberStyles.Float, CultureInfo.InvariantCulture, out size))
            {
                int hundredths = (int)Math.Round(size * 100.0);
                def.SetAttributeValue("sz", hundredths.ToString(CultureInfo.InvariantCulture));
            }

            string color = NormColor(style.TickColor);
            if (!string.IsNullOrEmpty(color))
            {
                def.Add(new XElement(A + "solidFill", new XElement(A + "srgbClr", new XAttribute("val", color))));
            }

            if (!string.IsNullOrEmpty(style.TickFont))
            {
                def.Add(new XElement(A + "latin", new XAttribute("typeface", style.TickFont)));
                def.Add(new XElement(A + "ea", new XAttribute("typeface", style.TickFont)));
            }

            var txPr = new XElement(
                C + "txPr",
                new XElement(A + "bodyPr"),
                new XElement(A + "lstStyle"),
                new XElement(A + "p", new XElement(A + "pPr", def), new XElement(A + "endParaRPr")));
            XElement cross = axis.Element(C + "crossAx");
            if (cross != null)
            {
                cross.AddBeforeSelf(txPr);
            }
            else
            {
                axis.Add(txPr);
            }
        }

        private static void WriteLn(XElement owner, string hex, string weight)
        {
            if (owner == null)
            {
                return;
            }

            XElement spPr = owner.Element(C + "spPr");
            if (spPr == null && owner.Name == C + "spPr")
            {
                spPr = owner;
            }

            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                owner.Add(spPr);
            }

            XElement ln = spPr.Element(A + "ln");
            if (ln == null)
            {
                ln = new XElement(A + "ln");
                spPr.Add(ln);
            }

            double pt;
            if (!string.IsNullOrEmpty(weight)
                && double.TryParse(weight, NumberStyles.Float, CultureInfo.InvariantCulture, out pt))
            {
                int emu = (int)Math.Round(pt * 12700.0);
                ln.SetAttributeValue("w", emu.ToString(CultureInfo.InvariantCulture));
            }

            if (string.IsNullOrEmpty(hex))
            {
                return;
            }

            XElement solid = ln.Element(A + "solidFill");
            if (solid == null)
            {
                solid = new XElement(A + "solidFill");
                ln.Add(solid);
            }

            XElement srgb = solid.Element(A + "srgbClr");
            if (srgb == null)
            {
                srgb = new XElement(A + "srgbClr");
                solid.Add(srgb);
            }

            srgb.SetAttributeValue("val", hex);
        }

        private static string ReadLegend(XDocument chartDoc)
        {
            XElement chart = chartDoc?.Descendants(C + "chart").FirstOrDefault();
            XElement legend = chart?.Elements(C + "legend").FirstOrDefault();
            if (chart != null && legend == null)
            {
                return "none";
            }

            string pos = legend?.Element(C + "legendPos")?.Attribute("val")?.Value;
            return LegendName(pos);
        }

        private static void ReadPrimaryScale(XDocument chartDoc, SeriesGrid grid)
        {
            XElement axis = FindValAx(chartDoc, false);
            if (axis == null || grid == null)
            {
                return;
            }

            XElement scaling = axis.Element(C + "scaling");
            grid.YMin = ReadDouble(scaling?.Element(C + "min"));
            grid.YMax = ReadDouble(scaling?.Element(C + "max"));
            grid.YMajor = ReadDouble(axis.Element(C + "majorUnit"));
        }

        private static string ReadValAxisTitle(XDocument chartDoc, bool secondary)
        {
            XElement axis = FindValAx(chartDoc, secondary);
            XElement text = axis?.Element(C + "title")?.Descendants(A + "t").FirstOrDefault();
            return text != null ? text.Value : null;
        }

        private static AxisYInfo StyleFromAxisX(AxisXInfo axis)
        {
            if (axis == null)
            {
                return null;
            }

            string format = axis.Format;
            if (string.Equals(format, "General", StringComparison.OrdinalIgnoreCase))
            {
                format = "@";
            }

            bool any = !string.IsNullOrEmpty(axis.Visible)
                || !string.IsNullOrEmpty(axis.TickFont)
                || !string.IsNullOrEmpty(axis.TickColor)
                || !string.IsNullOrEmpty(axis.TickSize)
                || !string.IsNullOrEmpty(axis.TickPosition)
                || !string.IsNullOrEmpty(axis.MajorTick)
                || !string.IsNullOrEmpty(axis.MinorTick)
                || !string.IsNullOrEmpty(axis.Grid)
                || !string.IsNullOrEmpty(format)
                || !string.IsNullOrEmpty(axis.Line)
                || !string.IsNullOrEmpty(axis.LineWeight)
                || !string.IsNullOrEmpty(axis.GridColor);
            if (!any)
            {
                return null;
            }

            return new AxisYInfo
            {
                Visible = axis.Visible,
                TickFont = axis.TickFont,
                TickColor = axis.TickColor,
                TickSize = axis.TickSize,
                TickPosition = axis.TickPosition,
                MajorTick = axis.MajorTick,
                MinorTick = axis.MinorTick,
                Grid = axis.Grid,
                Format = format,
                Line = axis.Line,
                LineWeight = axis.LineWeight,
                GridColor = axis.GridColor
            };
        }

        private static AxisXInfo ReadCatAxis(XDocument chartDoc)
        {
            XElement plot = chartDoc?.Descendants(C + "plotArea").FirstOrDefault();
            XElement cat = plot?.Elements(C + "catAx").FirstOrDefault();
            AxisYInfo style = ReadAxisChrome(cat);
            if (style == null)
            {
                return null;
            }

            return new AxisXInfo
            {
                Title = style.Title,
                Visible = style.Visible,
                TickFont = style.TickFont,
                TickColor = style.TickColor,
                TickSize = style.TickSize,
                TickPosition = style.TickPosition,
                MajorTick = style.MajorTick,
                MinorTick = style.MinorTick,
                Grid = style.Grid,
                Format = style.Format,
                Line = style.Line,
                LineWeight = style.LineWeight,
                GridColor = style.GridColor
            };
        }

        private static string ReadCatAxisTitle(XDocument chartDoc)
        {
            XElement plot = chartDoc?.Descendants(C + "plotArea").FirstOrDefault();
            XElement cat = plot?.Elements(C + "catAx").FirstOrDefault();
            XElement text = cat?.Element(C + "title")?.Descendants(A + "t").FirstOrDefault();
            return text != null ? text.Value : null;
        }

        private static string ReadGridlines(XDocument chartDoc)
        {
            XElement axis = FindValAx(chartDoc, false);
            if (axis == null)
            {
                return null;
            }

            return axis.Element(C + "majorGridlines") != null ? "true" : "false";
        }

        private static readonly XNamespace Mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";
        private static readonly XNamespace C14 = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";

        private static int ReadChartStyle(XDocument chartDoc)
        {
            if (chartDoc?.Root == null)
            {
                return 0;
            }

            XElement newer = chartDoc.Root.Descendants(C14 + "style").FirstOrDefault();
            int style;
            if (TryStyleVal(newer, out style))
            {
                return style;
            }

            XElement plain = chartDoc.Root.Element(C + "style")
                ?? chartDoc.Root.Descendants(C + "style").FirstOrDefault();
            return TryStyleVal(plain, out style) ? style : 0;
        }

        private static bool TryStyleVal(XElement el, out int style)
        {
            style = 0;
            return el != null
                && int.TryParse(el.Attribute("val")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out style)
                && style > 0;
        }

        private static void WriteGridlines(XElement axis, string gridlines)
        {
            if (axis == null || string.IsNullOrEmpty(gridlines))
            {
                return;
            }

            if (IsFalse(gridlines))
            {
                axis.Element(C + "majorGridlines")?.Remove();
                return;
            }

            if (IsTrue(gridlines) && axis.Element(C + "majorGridlines") == null)
            {
                axis.Add(new XElement(C + "majorGridlines"));
            }
        }

        private static void WriteBarGap(XDocument chartDoc, int? gap, int? overlap)
        {
            XElement bar = chartDoc?.Descendants(C + "barChart").FirstOrDefault();
            if (bar == null || (!gap.HasValue && !overlap.HasValue))
            {
                return;
            }

            XElement anchor = bar.Element(C + "axId");
            if (gap.HasValue)
            {
                UpsertBefore(bar, anchor, "gapWidth", gap.Value.ToString(CultureInfo.InvariantCulture));
            }

            if (overlap.HasValue)
            {
                UpsertBefore(bar, anchor, "overlap", overlap.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static void UpsertBefore(XElement parent, XElement before, string name, string value)
        {
            XElement child = parent.Element(C + name);
            if (child == null)
            {
                child = new XElement(C + name);
                if (before != null)
                {
                    before.AddBeforeSelf(child);
                }
                else
                {
                    parent.Add(child);
                }
            }

            child.SetAttributeValue("val", value ?? "");
        }

        private static int? ReadIntVal(XElement el)
        {
            int n;
            string val = el?.Attribute("val")?.Value;
            return int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? (int?)n : null;
        }

        private static void WriteChartStyle(XDocument chartDoc, int style)
        {
            if (chartDoc?.Root == null || style <= 0)
            {
                return;
            }

            bool inC14 = style >= 101 && style <= 148;
            bool inClassic = style >= 1 && style <= 48;
            if (!inC14 && !inClassic)
            {
                return;
            }

            chartDoc.Root.Element(C + "style")?.Remove();
            foreach (XElement old in chartDoc.Root.Elements(Mc + "AlternateContent").ToList())
            {
                if (old.Descendants(C14 + "style").Any() || old.Descendants(C + "style").Any())
                {
                    old.Remove();
                }
            }

            int fallback = style >= 1 && style <= 48 ? style : 2;
            var choice = new XElement(
                Mc + "Choice",
                new XAttribute("Requires", "c14"),
                new XAttribute(XNamespace.Xmlns + "c14", C14.NamespaceName),
                new XElement(C14 + "style", new XAttribute("val", style.ToString(CultureInfo.InvariantCulture))));
            var fb = new XElement(
                Mc + "Fallback",
                new XElement(C + "style", new XAttribute("val", fallback.ToString(CultureInfo.InvariantCulture))));
            var alt = new XElement(
                Mc + "AlternateContent",
                new XAttribute(XNamespace.Xmlns + "mc", Mc.NamespaceName),
                choice,
                fb);
            XElement chart = chartDoc.Root.Element(C + "chart");
            if (chart != null)
            {
                chart.AddBeforeSelf(alt);
            }
            else
            {
                chartDoc.Root.Add(alt);
            }
        }

        private static AxisYInfo ReadValAxisChrome(XDocument chartDoc)
        {
            return ReadAxisChrome(FindValAx(chartDoc, true));
        }

        private static AxisYInfo ReadAxisChrome(XElement axis)
        {
            if (axis == null)
            {
                return null;
            }

            var style = new AxisYInfo();
            string title = axis.Element(C + "title")?.Descendants(A + "t").FirstOrDefault()?.Value;
            if (!string.IsNullOrEmpty(title))
            {
                style.Title = title;
            }

            string tickPos = axis.Element(C + "tickLblPos")?.Attribute("val")?.Value;
            if (!string.IsNullOrEmpty(tickPos))
            {
                style.TickPosition = TickPosFromOoxml(tickPos);
                style.Visible = string.Equals(tickPos, "none", StringComparison.OrdinalIgnoreCase) ? "false" : "true";
            }

            string major = axis.Element(C + "majorTickMark")?.Attribute("val")?.Value;
            if (!string.IsNullOrEmpty(major))
            {
                style.MajorTick = TickMarkFromOoxml(major);
            }

            string minor = axis.Element(C + "minorTickMark")?.Attribute("val")?.Value;
            if (!string.IsNullOrEmpty(minor))
            {
                style.MinorTick = TickMarkFromOoxml(minor);
            }

            string format = axis.Element(C + "numFmt")?.Attribute("formatCode")?.Value;
            if (!string.IsNullOrEmpty(format))
            {
                style.Format = format;
            }

            XElement grid = axis.Element(C + "majorGridlines");
            if (grid != null)
            {
                style.Grid = "true";
                style.GridColor = HashColor(ReadLnColor(grid));
            }

            style.Line = HashColor(ReadLnColor(axis));
            style.LineWeight = ReadLnWeight(axis);
            XElement def = axis.Element(C + "txPr")?.Descendants(A + "defRPr").FirstOrDefault();
            if (def != null)
            {
                XElement latin = def.Element(A + "latin") ?? def.Element(A + "ea");
                if (latin != null)
                {
                    style.TickFont = latin.Attribute("typeface")?.Value;
                }

                string sz = def.Attribute("sz")?.Value;
                double hundredths;
                if (!string.IsNullOrEmpty(sz)
                    && double.TryParse(sz, NumberStyles.Float, CultureInfo.InvariantCulture, out hundredths))
                {
                    style.TickSize = (hundredths / 100.0).ToString(CultureInfo.InvariantCulture);
                }

                XElement tickFill = def.Element(A + "solidFill")?.Element(A + "srgbClr");
                style.TickColor = HashColor(tickFill?.Attribute("val")?.Value);
            }

            if (string.IsNullOrEmpty(style.Title)
                && string.IsNullOrEmpty(style.Visible)
                && string.IsNullOrEmpty(style.TickFont)
                && string.IsNullOrEmpty(style.TickColor)
                && string.IsNullOrEmpty(style.Format)
                && string.IsNullOrEmpty(style.Line)
                && string.IsNullOrEmpty(style.Grid))
            {
                return null;
            }

            return style;
        }

        private static void WriteSeriesNoFill(XElement ser)
        {
            if (ser == null)
            {
                return;
            }

            XElement spPr = ser.Element(C + "spPr");
            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                XElement ext = ser.Element(C + "extLst");
                if (ext != null)
                {
                    ext.AddBeforeSelf(spPr);
                }
                else
                {
                    ser.Add(spPr);
                }
            }

            spPr.Element(A + "solidFill")?.Remove();
            spPr.Element(A + "gradFill")?.Remove();
            if (spPr.Element(A + "noFill") == null)
            {
                spPr.AddFirst(new XElement(A + "noFill"));
            }
        }

        private static string ReadFillColor(XElement ser)
        {
            XElement spPr = ser?.Element(C + "spPr");
            if (spPr == null)
            {
                return null;
            }

            if (spPr.Element(A + "noFill") != null)
            {
                return "none";
            }

            return spPr.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        }

        private static string ReadLineColor(XElement ser)
        {
            return ReadLnColor(ser);
        }

        private static string ReadLineWeight(XElement ser)
        {
            return ReadLnWeight(ser);
        }

        private static void WriteMarker(XElement ser, SeriesCol col)
        {
            if (ser == null || col == null || IsNone(col.Marker))
            {
                return;
            }

            string fill = NormColor(col.MarkerFill);
            string line = NormColor(col.MarkerColor);
            int size;
            bool hasSize = int.TryParse(col.MarkerSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0;
            bool hasSymbol = !string.IsNullOrEmpty(col.Marker);
            if (!hasSymbol && !hasSize && string.IsNullOrEmpty(fill) && string.IsNullOrEmpty(line))
            {
                return;
            }

            XElement marker = ser.Element(C + "marker");
            if (marker == null)
            {
                marker = new XElement(C + "marker");
                ser.Add(marker);
            }

            if (hasSymbol)
            {
                SetVal(marker, "symbol", col.Marker.Trim());
            }

            if (hasSize)
            {
                SetVal(marker, "size", size.ToString(CultureInfo.InvariantCulture));
            }

            if (!string.IsNullOrEmpty(fill))
            {
                WriteMarkerFill(marker, fill);
            }

            if (!string.IsNullOrEmpty(line))
            {
                WriteLn(marker, line, null);
            }
        }

        private static void WriteMarkerFill(XElement marker, string hex)
        {
            XElement spPr = marker.Element(C + "spPr");
            if (spPr == null)
            {
                spPr = new XElement(C + "spPr");
                marker.Add(spPr);
            }

            spPr.Element(A + "noFill")?.Remove();
            spPr.Element(A + "gradFill")?.Remove();
            XElement solid = spPr.Element(A + "solidFill");
            if (solid == null)
            {
                solid = new XElement(A + "solidFill");
                XElement ln = spPr.Element(A + "ln");
                if (ln != null)
                {
                    ln.AddBeforeSelf(solid);
                }
                else
                {
                    spPr.Add(solid);
                }
            }

            XElement srgb = solid.Element(A + "srgbClr");
            if (srgb == null)
            {
                srgb = new XElement(A + "srgbClr");
                solid.Add(srgb);
            }

            srgb.SetAttributeValue("val", hex);
        }

        private static string ReadMarker(XElement ser)
        {
            return ser?.Element(C + "marker")?.Element(C + "symbol")?.Attribute("val")?.Value;
        }

        private static string ReadMarkerSize(XElement ser)
        {
            string val = ser?.Element(C + "marker")?.Element(C + "size")?.Attribute("val")?.Value;
            int size;
            if (int.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0)
            {
                return size.ToString(CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static string ReadMarkerColor(XElement ser)
        {
            return ser?.Element(C + "marker")
                ?.Element(C + "spPr")
                ?.Element(A + "ln")
                ?.Element(A + "solidFill")
                ?.Element(A + "srgbClr")
                ?.Attribute("val")?.Value;
        }

        private static string ReadMarkerFill(XElement ser)
        {
            return ser?.Element(C + "marker")
                ?.Element(C + "spPr")
                ?.Element(A + "solidFill")
                ?.Element(A + "srgbClr")
                ?.Attribute("val")?.Value;
        }

        private static string LabelPosToOoxml(string name)
        {
            switch ((name ?? "").Trim().ToLowerInvariant())
            {
                case "outside":
                    return "outEnd";
                case "above":
                    return "t";
                case "below":
                    return "b";
                case "center":
                    return "ctr";
                case "inside":
                    return "inEnd";
                default:
                    return null;
            }
        }

        private static XElement LabelDefRPr(XElement ser)
        {
            return ser?.Element(C + "dLbls")?.Element(C + "txPr")?.Descendants(A + "defRPr").FirstOrDefault();
        }

        private static string ReadLabelFont(XElement ser)
        {
            XElement latin = LabelDefRPr(ser)?.Element(A + "latin");
            string face = latin?.Attribute("typeface")?.Value;
            return string.IsNullOrEmpty(face) ? null : face;
        }

        private static string ReadLabelSize(XElement ser)
        {
            string sz = LabelDefRPr(ser)?.Attribute("sz")?.Value;
            int hundredths;
            if (!int.TryParse(sz, NumberStyles.Integer, CultureInfo.InvariantCulture, out hundredths) || hundredths <= 0)
            {
                return null;
            }

            return (hundredths / 100.0).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static string ReadLabelColor(XElement ser)
        {
            return LabelDefRPr(ser)?.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        }

        private static string ReadLabelFormat(XElement ser)
        {
            return ser?.Element(C + "dLbls")?.Element(C + "numFmt")?.Attribute("formatCode")?.Value;
        }

        private static string ReadGradient(XElement ser)
        {
            XElement gsLst = ser?.Element(C + "spPr")?.Element(A + "gradFill")?.Element(A + "gsLst");
            if (gsLst == null)
            {
                return null;
            }

            var sb = new StringBuilder();
            foreach (XElement gs in gsLst.Elements(A + "gs"))
            {
                int posRaw;
                if (!int.TryParse(gs.Attribute("pos")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out posRaw))
                {
                    continue;
                }

                string hex = gs.Element(A + "srgbClr")?.Attribute("val")?.Value;
                if (string.IsNullOrEmpty(hex))
                {
                    continue;
                }

                int alpha = 100000;
                string alphaRaw = gs.Element(A + "srgbClr")?.Element(A + "alpha")?.Attribute("val")?.Value;
                int parsedAlpha;
                if (int.TryParse(alphaRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out parsedAlpha))
                {
                    alpha = parsedAlpha;
                }

                if (sb.Length > 0)
                {
                    sb.Append(';');
                }

                double pos = posRaw / 100000.0;
                double transparency = 1.0 - (alpha / 100000.0);
                sb.Append(pos.ToString("0.##", CultureInfo.InvariantCulture))
                    .Append(":#")
                    .Append(hex.TrimStart('#').ToUpperInvariant())
                    .Append('@')
                    .Append(transparency.ToString("0.##", CultureInfo.InvariantCulture));
            }

            return sb.Length == 0 ? null : sb.ToString();
        }

        private static string ReadGradientAngle(XElement ser)
        {
            string ang = ser?.Element(C + "spPr")?.Element(A + "gradFill")?.Element(A + "lin")?.Attribute("ang")?.Value;
            int raw;
            if (!int.TryParse(ang, NumberStyles.Integer, CultureInfo.InvariantCulture, out raw))
            {
                return null;
            }

            return (raw / 60000.0).ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static string ReadLabelPos(XElement ser)
        {
            string pos = ser?.Element(C + "dLbls")?.Element(C + "dLblPos")?.Attribute("val")?.Value;
            switch ((pos ?? "").Trim())
            {
                case "outEnd":
                    return "outside";
                case "t":
                    return "above";
                case "b":
                    return "below";
                case "ctr":
                    return "center";
                case "inEnd":
                    return "inside";
                default:
                    return null;
            }
        }

        private static bool? ReadShowDataLabels(XElement ser)
        {
            XElement labels = ser?.Element(C + "dLbls");
            if (labels == null)
            {
                return null;
            }

            string deleted = labels.Element(C + "delete")?.Attribute("val")?.Value;
            if (deleted == "1" || string.Equals(deleted, "true", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return true;
        }

        private static bool? ReadShowPercent(XElement ser)
        {
            string val = ser?.Element(C + "dLbls")?.Element(C + "showPercent")?.Attribute("val")?.Value;
            if (val == "1" || string.Equals(val, "true", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (val == "0" || string.Equals(val, "false", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return null;
        }

        private static bool? ReadShowVal(XElement ser)
        {
            string val = ser?.Element(C + "dLbls")?.Element(C + "showVal")?.Attribute("val")?.Value;
            if (val == "1" || string.Equals(val, "true", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (val == "0" || string.Equals(val, "false", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return null;
        }

        private static string ReadLnColor(XElement owner)
        {
            if (owner == null)
            {
                return null;
            }

            XElement spPr = owner.Element(C + "spPr") ?? (owner.Name == C + "spPr" ? owner : null);
            return spPr?.Element(A + "ln")?.Element(A + "solidFill")?.Element(A + "srgbClr")?.Attribute("val")?.Value;
        }

        private static string ReadLnWeight(XElement owner)
        {
            XElement spPr = owner?.Element(C + "spPr") ?? (owner != null && owner.Name == C + "spPr" ? owner : null);
            string w = spPr?.Element(A + "ln")?.Attribute("w")?.Value;
            double emu;
            if (string.IsNullOrEmpty(w) || !double.TryParse(w, NumberStyles.Float, CultureInfo.InvariantCulture, out emu))
            {
                return null;
            }

            return (emu / 12700.0).ToString(CultureInfo.InvariantCulture);
        }

        private static XElement FindValAx(XDocument chartDoc, bool secondary)
        {
            XElement plot = chartDoc?.Descendants(C + "plotArea").FirstOrDefault();
            if (plot == null)
            {
                return null;
            }

            XElement fallback = null;
            foreach (XElement val in plot.Elements(C + "valAx"))
            {
                string pos = val.Element(C + "axPos")?.Attribute("val")?.Value;
                bool right = string.Equals(pos, "r", StringComparison.OrdinalIgnoreCase);
                if (secondary)
                {
                    if (right)
                    {
                        return val;
                    }

                    continue;
                }

                if (right)
                {
                    continue;
                }

                if (string.Equals(pos, "l", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(pos, "b", StringComparison.OrdinalIgnoreCase))
                {
                    return val;
                }

                if (fallback == null)
                {
                    fallback = val;
                }
            }

            return secondary ? null : fallback;
        }

        private static void SetVal(XElement parent, string name, string value)
        {
            if (parent == null || string.IsNullOrEmpty(name))
            {
                return;
            }

            XElement child = parent.Element(C + name);
            if (child == null)
            {
                child = new XElement(C + name);
                parent.Add(child);
            }

            child.SetAttributeValue("val", value ?? "");
        }

        private static double? ReadDouble(XElement el)
        {
            double n;
            string text = el?.Attribute("val")?.Value;
            if (string.IsNullOrEmpty(text) || !double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out n))
            {
                return null;
            }

            return n;
        }

        private static string HashColor(string hex)
        {
            if (string.IsNullOrEmpty(hex))
            {
                return null;
            }

            return hex.StartsWith("#", StringComparison.Ordinal) ? hex : "#" + hex;
        }

        private static bool IsNone(string value)
        {
            return string.Equals(value, "none", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "1", StringComparison.Ordinal);
        }

        private static bool IsFalse(string value)
        {
            return string.Equals(value, "false", StringComparison.OrdinalIgnoreCase)
                || string.Equals(value, "0", StringComparison.Ordinal);
        }

        private static string LegendPos(string legend)
        {
            if (string.IsNullOrEmpty(legend))
            {
                return null;
            }

            switch (legend.Trim().ToLowerInvariant())
            {
                case "bottom":
                case "b":
                    return "b";
                case "top":
                case "t":
                    return "t";
                case "left":
                case "l":
                    return "l";
                case "right":
                case "r":
                    return "r";
                default:
                    return null;
            }
        }

        private static string LegendName(string pos)
        {
            switch ((pos ?? "").Trim().ToLowerInvariant())
            {
                case "b":
                    return "bottom";
                case "t":
                    return "top";
                case "l":
                    return "left";
                case "r":
                    return "right";
                default:
                    return null;
            }
        }

        private static string TickPosToOoxml(string pos)
        {
            switch ((pos ?? "").Trim().ToLowerInvariant())
            {
                case "next":
                case "nextto":
                    return "nextTo";
                case "low":
                    return "low";
                case "high":
                    return "high";
                case "none":
                    return "none";
                default:
                    return null;
            }
        }

        private static string TickPosFromOoxml(string pos)
        {
            switch ((pos ?? "").Trim().ToLowerInvariant())
            {
                case "nextto":
                    return "next";
                case "low":
                    return "low";
                case "high":
                    return "high";
                case "none":
                    return "none";
                default:
                    return pos;
            }
        }

        private static string TickMarkToOoxml(string mark)
        {
            switch ((mark ?? "").Trim().ToLowerInvariant())
            {
                case "outside":
                case "out":
                    return "out";
                case "inside":
                case "in":
                    return "in";
                case "none":
                    return "none";
                case "cross":
                    return "cross";
                default:
                    return null;
            }
        }

        private static string TickMarkFromOoxml(string mark)
        {
            switch ((mark ?? "").Trim().ToLowerInvariant())
            {
                case "out":
                case "outside":
                    return "outside";
                case "in":
                case "inside":
                    return "inside";
                case "none":
                    return "none";
                case "cross":
                    return "cross";
                default:
                    return mark;
            }
        }
    }
}
