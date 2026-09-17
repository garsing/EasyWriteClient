using System;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using WordAddIn1.PresentationHost;

namespace PptChartRoundtripTest
{
    internal static class ChartAttrMatch
    {
        public static string Check(IList<AttrCheck> checks, PptHtmlChartReadModel model)
        {
            if (checks == null || checks.Count == 0)
            {
                return null;
            }

            var fails = new List<string>();
            for (int i = 0; i < checks.Count; i++)
            {
                AttrCheck one = checks[i];
                if (one == null || string.Equals(one.Mode, "applyonly", StringComparison.Ordinal))
                {
                    continue;
                }

                string actual = Lookup(model, one.Key);
                if (!Same(one, actual))
                {
                    fails.Add(one.Key + " 期望 " + (one.Expected ?? "(null)")
                        + " 实际 " + (actual ?? "(null)"));
                }
            }

            if (fails.Count == 0)
            {
                return null;
            }

            int n = Math.Min(4, fails.Count);
            return string.Join("；", fails.GetRange(0, n));
        }

        private static string Lookup(PptHtmlChartReadModel model, string key)
        {
            if (model == null || string.IsNullOrEmpty(key))
            {
                return null;
            }

            if (key.StartsWith("S", StringComparison.Ordinal) && key.IndexOf('.') > 0)
            {
                int dot = key.IndexOf('.');
                if (int.TryParse(key.Substring(1, dot - 1), NumberStyles.Integer, CultureInfo.InvariantCulture, out int idx))
                {
                    return Prop(ChartHtml.ValueCol(model.Grid, idx), key.Substring(dot + 1));
                }
            }

            PptHtmlChartFormat fmt = model.Format;
            if (key.StartsWith("AxisXStyle.", StringComparison.Ordinal))
            {
                return Prop(fmt == null ? null : fmt.AxisXStyle, key.Substring("AxisXStyle.".Length));
            }

            if (key.StartsWith("AxisYStyle.", StringComparison.Ordinal))
            {
                return Prop(fmt == null ? null : fmt.AxisYStyle, key.Substring("AxisYStyle.".Length));
            }

            if (key.StartsWith("AxisY2Style.", StringComparison.Ordinal))
            {
                return Prop(fmt == null ? null : fmt.AxisY2Style, key.Substring("AxisY2Style.".Length));
            }

            return Prop(fmt, key);
        }

        private static string Prop(object target, string name)
        {
            if (target == null || string.IsNullOrEmpty(name))
            {
                return null;
            }

            PropertyInfo p = target.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public);
            if (p == null)
            {
                return null;
            }

            object v = p.GetValue(target, null);
            return v == null ? null : Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        private static bool Same(AttrCheck one, string actual)
        {
            string mode = one.Mode ?? "exact";
            string expected = one.Expected;
            if (mode == "hex")
            {
                return string.Equals(NormHex(expected), NormHex(actual), StringComparison.Ordinal);
            }

            if (mode == "num")
            {
                return CloseNum(expected, actual, 0.06);
            }

            if (mode == "box")
            {
                return CloseBox(expected, actual, 4.0);
            }

            return string.Equals(expected ?? "", actual ?? "", StringComparison.Ordinal);
        }

        private static string NormHex(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return raw;
            }

            string t = raw.Trim();
            if (string.Equals(t, "none", StringComparison.OrdinalIgnoreCase))
            {
                return "none";
            }

            if (t.StartsWith("#", StringComparison.Ordinal))
            {
                t = t.Substring(1);
            }

            return t.ToUpperInvariant();
        }

        private static bool CloseNum(string expected, string actual, double eps)
        {
            return double.TryParse(expected, NumberStyles.Float, CultureInfo.InvariantCulture, out double e)
                && double.TryParse(actual, NumberStyles.Float, CultureInfo.InvariantCulture, out double a)
                && Math.Abs(e - a) <= eps;
        }

        private static bool CloseBox(string expected, string actual, double eps)
        {
            if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(actual))
            {
                return false;
            }

            string[] a = expected.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            string[] b = actual.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (a.Length < 4 || b.Length < 4)
            {
                return false;
            }

            for (int i = 0; i < 4; i++)
            {
                if (!double.TryParse(a[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double ea)
                    || !double.TryParse(b[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double eb)
                    || Math.Abs(ea - eb) > eps)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
