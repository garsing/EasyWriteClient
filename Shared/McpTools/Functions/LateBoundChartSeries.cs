using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;

namespace WordAddIn1
{
    /// <summary>
    /// WPP / WPS 文字共用的晚绑定 Chart 系列读取（对齐 PptHtmlChartIo.TryReadGridFromSeries）。
    /// 不走 Word.InlineShape / 内嵌 Excel Chart COM。
    /// </summary>
    internal static class LateBoundChartSeries
    {
        internal sealed class Grid
        {
            public List<string> Categories { get; set; }

            public List<Col> Series { get; set; }
        }

        internal sealed class Col
        {
            public string Name { get; set; }

            public string Color { get; set; }

            public List<string> Values { get; set; }
        }

        public static bool TryRead(object chart, out Grid grid, out string error)
        {
            grid = null;
            error = null;
            if (chart == null)
            {
                error = "图表没有系列";
                return false;
            }

            try
            {
                object sc = TryInvoke(chart, "SeriesCollection");
                if (sc == null)
                {
                    sc = GetProp(chart, "SeriesCollection");
                }

                int count = Convert.ToInt32(GetProp(sc, "Count"));
                if (count < 1)
                {
                    error = "图表没有系列";
                    return false;
                }

                object s1 = GetSeries(chart, 1);
                List<string> cats = ToStringList(GetProp(s1, "XValues"));
                if (cats.Count == 0)
                {
                    cats = TryReadCategoryAxisNames(chart);
                }

                if (cats.Count == 0)
                {
                    error = "图表没有类别";
                    return false;
                }

                var series = new List<Col>();
                for (int i = 1; i <= count; i++)
                {
                    object one = GetSeries(chart, i);
                    string name = Convert.ToString(GetProp(one, "Name") ?? ("系列" + i));
                    series.Add(new Col
                    {
                        Name = name,
                        Color = TryReadSeriesColor(one),
                        Values = ToStringList(GetProp(one, "Values"))
                    });
                }

                grid = new Grid
                {
                    Categories = cats,
                    Series = series
                };
                return true;
            }
            catch (Exception ex)
            {
                error = "读 SeriesCollection 失败: " + ex.Message;
                return false;
            }
        }

        private static object GetSeries(object chart, int index)
        {
            object series = TryInvoke(chart, "SeriesCollection", index);
            if (series != null)
            {
                return series;
            }

            object sc = TryInvoke(chart, "SeriesCollection");
            if (sc == null)
            {
                sc = GetProp(chart, "SeriesCollection");
            }

            return GetIndexed(sc, index);
        }

        private static List<string> TryReadCategoryAxisNames(object chart)
        {
            var cats = new List<string>();
            try
            {
                object axis = TryInvoke(chart, "Axes", 1, 1)
                    ?? TryInvoke(chart, "Axes", 1);
                if (axis == null)
                {
                    return cats;
                }

                return ToStringList(GetProp(axis, "CategoryNames"));
            }
            catch
            {
                return cats;
            }
        }

        private static string TryReadSeriesColor(object series)
        {
            try
            {
                object fmt = GetProp(series, "Format");
                object fill = GetProp(fmt, "Fill");
                object vis = fill == null ? null : GetProp(fill, "Visible");
                if (vis != null && Convert.ToInt32(vis) == 0)
                {
                    return "none";
                }

                object fc = GetProp(fill, "ForeColor");
                object rgb = GetProp(fc, "RGB");
                if (rgb != null)
                {
                    int office = Convert.ToInt32(Convert.ToDouble(rgb));
                    int r = office & 0xFF;
                    int g = (office >> 8) & 0xFF;
                    int b = (office >> 16) & 0xFF;
                    return "#" + r.ToString("X2", CultureInfo.InvariantCulture)
                        + g.ToString("X2", CultureInfo.InvariantCulture)
                        + b.ToString("X2", CultureInfo.InvariantCulture);
                }
            }
            catch
            {
            }

            return null;
        }

        private static List<string> ToStringList(object values)
        {
            var list = new List<string>();
            if (values == null)
            {
                return list;
            }

            if (values is Array arr)
            {
                foreach (object item in arr)
                {
                    list.Add(FormatCell(item));
                }

                return list;
            }

            list.Add(FormatCell(values));
            return list;
        }

        private static string FormatCell(object v)
        {
            if (v == null || v is DBNull)
            {
                return "";
            }

            if (v is double d)
            {
                return d.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is float f)
            {
                return f.ToString("0.####", CultureInfo.InvariantCulture);
            }

            if (v is decimal m)
            {
                return m.ToString("0.####", CultureInfo.InvariantCulture);
            }

            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? "";
        }

        private static object TryInvoke(object target, string name, params object[] args)
        {
            if (target == null)
            {
                return null;
            }

            try
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    target,
                    args);
            }
            catch
            {
                return null;
            }
        }

        private static object GetProp(object target, string name)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.Public,
                null,
                target,
                null);
        }

        private static object GetIndexed(object collection, int index)
        {
            if (collection == null)
            {
                return null;
            }

            try
            {
                return collection.GetType().InvokeMember(
                    "Item",
                    BindingFlags.GetProperty | BindingFlags.InvokeMethod | BindingFlags.Instance | BindingFlags.Public,
                    null,
                    collection,
                    new object[] { index });
            }
            catch
            {
                return null;
            }
        }
    }
}
