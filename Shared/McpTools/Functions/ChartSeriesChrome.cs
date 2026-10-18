using System;
using System.Globalization;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 系列线条、标记、标签、渐变，以及轴刻度字体。创建和抽图共用。
    /// </summary>
    internal static class ChartSeriesChrome
    {
        public static void ReadColumnAttrs(SeriesInfo series, System.Xml.Linq.XElement column)
        {
            if (series == null || column == null)
            {
                return;
            }

            series.Line = Attr(column, "line");
            series.LineWeight = Attr(column, "lineWeight");
            series.Marker = Attr(column, "marker");
            series.MarkerSize = Attr(column, "markerSize");
            series.MarkerColor = Attr(column, "markerColor");
            series.MarkerFill = Attr(column, "markerFill");
            series.LabelPosition = Attr(column, "labelPosition");
            series.LabelFont = Attr(column, "labelFont");
            series.LabelSize = Attr(column, "labelSize");
            series.LabelColor = Attr(column, "labelColor");
            series.LabelFormat = Attr(column, "labelFormat");
            series.FillGradient = Attr(column, "fillGradient");
            series.FillAngle = Attr(column, "fillAngle");
            if (bool.TryParse(Attr(column, "showValue"), out bool showValue))
            {
                series.ShowValue = showValue;
            }

            if (bool.TryParse(Attr(column, "showPercentage"), out bool showPercentage))
            {
                series.ShowPercentage = showPercentage;
            }
        }

        public static void WriteColumnAttrs(System.Xml.Linq.XElement col, SeriesInfo series)
        {
            if (col == null || series == null)
            {
                return;
            }

            Add(col, "line", series.Line);
            Add(col, "lineWeight", series.LineWeight);
            Add(col, "marker", series.Marker);
            Add(col, "markerSize", series.MarkerSize);
            Add(col, "markerColor", series.MarkerColor);
            Add(col, "markerFill", series.MarkerFill);
            Add(col, "labelPosition", series.LabelPosition);
            Add(col, "labelFont", series.LabelFont);
            Add(col, "labelSize", series.LabelSize);
            Add(col, "labelColor", series.LabelColor);
            Add(col, "labelFormat", series.LabelFormat);
            Add(col, "fillGradient", series.FillGradient);
            Add(col, "fillAngle", series.FillAngle);
            if (series.ShowValue.HasValue)
            {
                col.SetAttributeValue("showValue", series.ShowValue.Value ? "true" : "false");
            }

            if (series.ShowPercentage.HasValue)
            {
                col.SetAttributeValue("showPercentage", series.ShowPercentage.Value ? "true" : "false");
            }
        }

        public static void Apply(dynamic series, SeriesInfo info)
        {
            if (series == null || info == null)
            {
                return;
            }

            try
            {
                if (ChartEngine.TryParseOfficeBgr(info.Line, out int line))
                {
                    series.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    series.Format.Line.ForeColor.RGB = line;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (double.TryParse(info.LineWeight, NumberStyles.Float, CultureInfo.InvariantCulture, out double weight) && weight > 0)
                {
                    series.Format.Line.Weight = weight;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                int? marker = MarkerToXl(info.Marker);
                if (marker.HasValue)
                {
                    series.MarkerStyle = marker.Value;
                }

                if (int.TryParse(info.MarkerSize, NumberStyles.Integer, CultureInfo.InvariantCulture, out int size) && size > 0)
                {
                    series.MarkerSize = size;
                }

                if (ChartEngine.TryParseOfficeBgr(info.MarkerColor, out int markerLine))
                {
                    series.MarkerForegroundColor = markerLine;
                }

                if (ChartEngine.TryParseOfficeBgr(info.MarkerFill, out int markerFill))
                {
                    series.MarkerBackgroundColor = markerFill;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (!string.IsNullOrEmpty(info.LabelPosition)
                    || !string.IsNullOrEmpty(info.LabelFont)
                    || !string.IsNullOrEmpty(info.LabelColor)
                    || info.ShowValue.HasValue
                    || info.ShowPercentage.HasValue)
                {
                    series.HasDataLabels = true;
                }

                if (series.HasDataLabels)
                {
                    dynamic labels = series.DataLabels;
                    int? pos = LabelToXl(info.LabelPosition);
                    if (pos.HasValue)
                    {
                        labels.Position = pos.Value;
                    }

                    if (info.ShowValue.HasValue)
                    {
                        labels.ShowValue = info.ShowValue.Value;
                    }

                    if (info.ShowPercentage.HasValue)
                    {
                        labels.ShowPercentage = info.ShowPercentage.Value;
                    }

                    if (!string.IsNullOrEmpty(info.LabelFont))
                    {
                        labels.Font.Name = info.LabelFont;
                    }

                    if (double.TryParse(info.LabelSize, NumberStyles.Float, CultureInfo.InvariantCulture, out double labelSize) && labelSize > 0)
                    {
                        labels.Font.Size = labelSize;
                    }

                    if (ChartEngine.TryParseOfficeBgr(info.LabelColor, out int labelColor))
                    {
                        labels.Font.Color = labelColor;
                    }
                }
            }
            catch (Exception)
            {
            }

            ApplyGradient(series, info.FillGradient, info.FillAngle);
        }

        public static void Capture(dynamic series, SeriesInfo info)
        {
            if (series == null || info == null)
            {
                return;
            }

            try
            {
                int rgb = (int)series.Format.Line.ForeColor.RGB;
                if (rgb != 0)
                {
                    info.Line = Hex(rgb);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                double weight = Convert.ToDouble(series.Format.Line.Weight, CultureInfo.InvariantCulture);
                if (weight > 0)
                {
                    info.LineWeight = weight.ToString("0.##", CultureInfo.InvariantCulture);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                info.Marker = MarkerFromXl(Convert.ToInt32(series.MarkerStyle));
                int size = Convert.ToInt32(series.MarkerSize);
                if (size > 0 && info.Marker != null)
                {
                    info.MarkerSize = size.ToString(CultureInfo.InvariantCulture);
                }

                info.MarkerColor = Hex(Convert.ToInt32(series.MarkerForegroundColor));
                info.MarkerFill = Hex(Convert.ToInt32(series.MarkerBackgroundColor));
            }
            catch (Exception)
            {
            }

            try
            {
                if (series.HasDataLabels)
                {
                    dynamic labels = series.DataLabels;
                    info.LabelPosition = LabelFromXl(Convert.ToInt32(labels.Position));
                    try
                    {
                        info.ShowValue = (bool)labels.ShowValue;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        info.ShowPercentage = (bool)labels.ShowPercentage;
                    }
                    catch (Exception)
                    {
                    }

                    try
                    {
                        string font = Convert.ToString(labels.Font.Name);
                        if (!string.IsNullOrWhiteSpace(font))
                        {
                            info.LabelFont = font;
                        }

                        double labelSize = Convert.ToDouble(labels.Font.Size, CultureInfo.InvariantCulture);
                        if (labelSize > 0)
                        {
                            info.LabelSize = labelSize.ToString("0.##", CultureInfo.InvariantCulture);
                        }

                        info.LabelColor = Hex(Convert.ToInt32(labels.Font.Color));
                    }
                    catch (Exception)
                    {
                    }
                }
            }
            catch (Exception)
            {
            }

            CaptureGradient(series, info);
        }

        public static void ApplyAxisStroke(dynamic axis, string line, string weight, string gridColor)
        {
            if (axis == null)
            {
                return;
            }

            try
            {
                if (ChartEngine.TryParseOfficeBgr(line, out int lineBgr))
                {
                    axis.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
                    axis.Format.Line.ForeColor.RGB = lineBgr;
                }

                if (double.TryParse(weight, NumberStyles.Float, CultureInfo.InvariantCulture, out double w) && w > 0)
                {
                    axis.Format.Line.Weight = w;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (ChartEngine.TryParseOfficeBgr(gridColor, out int gridBgr))
                {
                    axis.HasMajorGridlines = true;
                    axis.MajorGridlines.Format.Line.ForeColor.RGB = gridBgr;
                }
            }
            catch (Exception)
            {
            }
        }

        public static void CaptureAxisStroke(dynamic axis, Action<string, string, string> assign)
        {
            if (axis == null || assign == null)
            {
                return;
            }

            string line = null;
            string weight = null;
            string gridColor = null;
            try
            {
                line = Hex(Convert.ToInt32(axis.Format.Line.ForeColor.RGB));
                double w = Convert.ToDouble(axis.Format.Line.Weight, CultureInfo.InvariantCulture);
                if (w > 0)
                {
                    weight = w.ToString("0.##", CultureInfo.InvariantCulture);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                gridColor = Hex(Convert.ToInt32(axis.MajorGridlines.Format.Line.ForeColor.RGB));
            }
            catch (Exception)
            {
            }

            assign(line, weight, gridColor);
        }

        public static void ApplyAxis(dynamic axis, string visible, string font, string color, string size, string format)
        {
            if (axis == null)
            {
                return;
            }

            try
            {
                if (string.Equals(visible, "false", StringComparison.OrdinalIgnoreCase))
                {
                    axis.TickLabelPosition = Word.XlTickLabelPosition.xlTickLabelPositionNone;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (!string.IsNullOrEmpty(font))
                {
                    axis.TickLabels.Font.Name = font;
                }

                if (double.TryParse(size, NumberStyles.Float, CultureInfo.InvariantCulture, out double n) && n > 0)
                {
                    axis.TickLabels.Font.Size = n;
                }

                if (ChartEngine.TryParseOfficeBgr(color, out int bgr))
                {
                    axis.TickLabels.Font.Color = bgr;
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (!string.IsNullOrEmpty(format))
                {
                    axis.TickLabels.NumberFormat = format;
                }
            }
            catch (Exception)
            {
            }
        }

        public static void ApplyTicks(dynamic axis, string tickPos, string major, string minor, string grid)
        {
            if (axis == null)
            {
                return;
            }

            try
            {
                if (!string.IsNullOrEmpty(tickPos))
                {
                    axis.TickLabelPosition = TickPosToXl(tickPos);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (!string.IsNullOrEmpty(major))
                {
                    axis.MajorTickMark = TickMarkToXl(major);
                }

                if (!string.IsNullOrEmpty(minor))
                {
                    axis.MinorTickMark = TickMarkToXl(minor);
                }
            }
            catch (Exception)
            {
            }

            try
            {
                if (string.Equals(grid, "true", StringComparison.OrdinalIgnoreCase))
                {
                    axis.HasMajorGridlines = true;
                }
                else if (string.Equals(grid, "false", StringComparison.OrdinalIgnoreCase))
                {
                    axis.HasMajorGridlines = false;
                }
            }
            catch (Exception)
            {
            }
        }

        public static void CaptureTicks(dynamic axis, Action<string, string, string, string> assign)
        {
            if (axis == null || assign == null)
            {
                return;
            }

            string pos = null;
            string major = null;
            string minor = null;
            string grid = null;
            try
            {
                pos = TickPosFromXl(Convert.ToInt32(axis.TickLabelPosition));
            }
            catch (Exception)
            {
            }

            try
            {
                major = TickMarkFromXl(Convert.ToInt32(axis.MajorTickMark));
                minor = TickMarkFromXl(Convert.ToInt32(axis.MinorTickMark));
            }
            catch (Exception)
            {
            }

            try
            {
                grid = axis.HasMajorGridlines ? "true" : "false";
            }
            catch (Exception)
            {
            }

            assign(pos, major, minor, grid);
        }

        private static int TickPosToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return -4142;
                case "low":
                    return -4134;
                case "high":
                    return -4127;
                default:
                    return 4;
            }
        }

        private static string TickPosFromXl(int pos)
        {
            if (pos == -4142)
            {
                return "none";
            }

            if (pos == -4134)
            {
                return "low";
            }

            if (pos == -4127)
            {
                return "high";
            }

            return "next";
        }

        private static int TickMarkToXl(string raw)
        {
            switch ((raw ?? "").Trim().ToLowerInvariant())
            {
                case "none":
                    return -4142;
                case "inside":
                    return 2;
                case "cross":
                    return 4;
                default:
                    return 3;
            }
        }

        private static string TickMarkFromXl(int mark)
        {
            if (mark == -4142)
            {
                return "none";
            }

            if (mark == 2)
            {
                return "inside";
            }

            if (mark == 4)
            {
                return "cross";
            }

            return "outside";
        }

        public static void CaptureAxis(dynamic axis, Action<string, string, string, string, string> assign)
        {
            if (axis == null || assign == null)
            {
                return;
            }

            string visible = "true";
            string font = null;
            string color = null;
            string size = null;
            string format = null;
            try
            {
                int pos = Convert.ToInt32(axis.TickLabelPosition);
                if (pos == (int)Word.XlTickLabelPosition.xlTickLabelPositionNone)
                {
                    visible = "false";
                }
            }
            catch (Exception)
            {
            }

            try
            {
                font = Convert.ToString(axis.TickLabels.Font.Name);
                double n = Convert.ToDouble(axis.TickLabels.Font.Size, CultureInfo.InvariantCulture);
                if (n > 0)
                {
                    size = n.ToString("0.##", CultureInfo.InvariantCulture);
                }

                color = Hex(Convert.ToInt32(axis.TickLabels.Font.Color));
            }
            catch (Exception)
            {
            }

            try
            {
                format = Convert.ToString(axis.TickLabels.NumberFormat);
            }
            catch (Exception)
            {
            }

            assign(visible, font, color, size, format);
        }

        private static void ApplyGradient(dynamic series, string gradient, string angle)
        {
            if (string.IsNullOrWhiteSpace(gradient))
            {
                return;
            }

            try
            {
                string[] parts = gradient.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 2)
                {
                    return;
                }

                if (!TryStopColor(parts[0], out int first) || !TryStopColor(parts[parts.Length - 1], out int last))
                {
                    return;
                }

                series.Format.Fill.TwoColorGradient(Microsoft.Office.Core.MsoGradientStyle.msoGradientVertical, 1);
                series.Format.Fill.ForeColor.RGB = first;
                series.Format.Fill.BackColor.RGB = last;
                if (double.TryParse(angle, NumberStyles.Float, CultureInfo.InvariantCulture, out double deg))
                {
                    series.Format.Fill.GradientAngle = (float)deg;
                }
            }
            catch (Exception)
            {
            }
        }

        private static void CaptureGradient(dynamic series, SeriesInfo info)
        {
            try
            {
                dynamic stops = series.Format.Fill.GradientStops;
                int count = (int)stops.Count;
                if (count < 2)
                {
                    return;
                }

                var sb = new StringBuilder();
                for (int i = 1; i <= count; i++)
                {
                    dynamic stop = stops.Item(i);
                    if (i > 1)
                    {
                        sb.Append(';');
                    }

                    double pos = Convert.ToDouble(stop.Position, CultureInfo.InvariantCulture);
                    double transparency = 0;
                    try
                    {
                        transparency = Convert.ToDouble(stop.Transparency, CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                    }

                    sb.Append(pos.ToString("0.##", CultureInfo.InvariantCulture))
                        .Append(':')
                        .Append(Hex(Convert.ToInt32(stop.Color.RGB)))
                        .Append('@')
                        .Append(transparency.ToString("0.##", CultureInfo.InvariantCulture));
                }

                info.FillGradient = sb.ToString();
                try
                {
                    double angle = Convert.ToDouble(series.Format.Fill.GradientAngle, CultureInfo.InvariantCulture);
                    info.FillAngle = angle.ToString("0.##", CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {
                }
            }
            catch (Exception)
            {
            }
        }

        private static bool TryStopColor(string part, out int bgr)
        {
            bgr = 0;
            int colon = part.IndexOf(':');
            int at = part.LastIndexOf('@');
            if (colon < 0)
            {
                return false;
            }

            string hex = at > colon ? part.Substring(colon + 1, at - colon - 1) : part.Substring(colon + 1);
            return ChartEngine.TryParseOfficeBgr(hex, out bgr);
        }

        private static int? MarkerToXl(string name)
        {
            switch ((name ?? "").Trim().ToLowerInvariant())
            {
                case "circle":
                    return 8;
                case "diamond":
                    return 2;
                case "square":
                    return 1;
                case "triangle":
                    return 3;
                case "star":
                    return 5;
                case "none":
                    return -4142;
                default:
                    return null;
            }
        }

        private static string MarkerFromXl(int value)
        {
            switch (value)
            {
                case 8:
                    return "circle";
                case 2:
                    return "diamond";
                case 1:
                    return "square";
                case 3:
                    return "triangle";
                case 5:
                    return "star";
                case -4142:
                    return "none";
                default:
                    return null;
            }
        }

        private static int? LabelToXl(string name)
        {
            switch ((name ?? "").Trim().ToLowerInvariant())
            {
                case "outside":
                    return 2;
                case "above":
                    return 0;
                case "center":
                    return -4108;
                case "inside":
                    return 3;
                default:
                    return null;
            }
        }

        private static string LabelFromXl(int value)
        {
            switch (value)
            {
                case 2:
                    return "outside";
                case 0:
                    return "above";
                case -4108:
                    return "center";
                case 3:
                    return "inside";
                default:
                    return null;
            }
        }

        public static string RgbHex(int rgbValue)
        {
            return Hex(rgbValue);
        }

        private static string Hex(int rgbValue)
        {
            int r = rgbValue & 0xFF;
            int g = (rgbValue >> 8) & 0xFF;
            int b = (rgbValue >> 16) & 0xFF;
            return string.Format(CultureInfo.InvariantCulture, "#{0:X2}{1:X2}{2:X2}", r, g, b);
        }

        private static string Attr(System.Xml.Linq.XElement column, string name)
        {
            return column.Attribute(name)?.Value;
        }

        private static void Add(System.Xml.Linq.XElement col, string name, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                col.SetAttributeValue(name, value);
            }
        }
    }
}
