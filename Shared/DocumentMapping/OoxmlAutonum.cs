using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml.Linq;

namespace WordAddIn1.DocumentMapping
{
    public sealed class OoxmlAutonum
    {
        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        private readonly Dictionary<string, XElement> _nums = new Dictionary<string, XElement>(StringComparer.Ordinal);
        private readonly Dictionary<string, XElement> _abstracts = new Dictionary<string, XElement>(StringComparer.Ordinal);
        private readonly Dictionary<string, XElement> _styles = new Dictionary<string, XElement>(StringComparer.Ordinal);
        private readonly Dictionary<string, int[]> _counters = new Dictionary<string, int[]>(StringComparer.Ordinal);

        public static OoxmlAutonum TryCreate(XDocument pkg)
        {
            try
            {
                return new OoxmlAutonum(pkg);
            }
            catch
            {
                return null;
            }
        }

        private OoxmlAutonum(XDocument pkg)
        {
            if (pkg == null)
            {
                return;
            }

            XElement numbering = FindPartRoot(pkg, "numbering.xml", W + "numbering");
            if (numbering != null)
            {
                foreach (XElement abs in numbering.Elements(W + "abstractNum"))
                {
                    string id = (string)abs.Attribute(W + "abstractNumId");
                    if (!string.IsNullOrEmpty(id))
                    {
                        _abstracts[id] = abs;
                    }
                }

                foreach (XElement num in numbering.Elements(W + "num"))
                {
                    string id = (string)num.Attribute(W + "numId");
                    if (!string.IsNullOrEmpty(id))
                    {
                        _nums[id] = num;
                    }
                }
            }

            XElement styles = FindPartRoot(pkg, "styles.xml", W + "styles");
            if (styles != null)
            {
                foreach (XElement st in styles.Elements(W + "style"))
                {
                    string id = (string)st.Attribute(W + "styleId");
                    if (!string.IsNullOrEmpty(id))
                    {
                        _styles[id] = st;
                    }
                }
            }
        }

        public string TryFormat(XElement paragraph)
        {
            try
            {
                if (paragraph == null || _nums.Count == 0)
                {
                    return null;
                }

                XElement numPr = FindNumPr(paragraph);
                if (numPr == null)
                {
                    return null;
                }

                string numId = Val(numPr.Element(W + "numId"));
                if (string.IsNullOrEmpty(numId) || numId == "0")
                {
                    return null;
                }

                int ilvl = 0;
                int.TryParse(Val(numPr.Element(W + "ilvl")) ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out ilvl);
                if (ilvl < 0)
                {
                    ilvl = 0;
                }

                XElement num;
                if (!_nums.TryGetValue(numId, out num))
                {
                    return null;
                }

                string absId = Val(num.Element(W + "abstractNumId"));
                XElement abs;
                if (string.IsNullOrEmpty(absId) || !_abstracts.TryGetValue(absId, out abs))
                {
                    return null;
                }

                XElement lvl = FindLvl(abs, num, ilvl);
                if (lvl == null)
                {
                    return null;
                }

                int count = Advance(numId, ilvl, StartOf(num, lvl, ilvl));
                string fmt = Val(lvl.Element(W + "numFmt")) ?? "decimal";
                string lvlText = Val(lvl.Element(W + "lvlText"));
                string token = FormatNumber(fmt, count, lvlText);
                if (string.IsNullOrEmpty(token))
                {
                    return null;
                }

                if (string.IsNullOrEmpty(lvlText))
                {
                    return token;
                }

                if (lvlText.IndexOf('%') < 0)
                {
                    return lvlText;
                }

                string visible = lvlText;
                for (int level = 0; level <= ilvl; level++)
                {
                    int n = level == ilvl ? count : Current(numId, level);
                    string piece = FormatNumber(fmt, n, null);
                    visible = visible.Replace("%" + (level + 1).ToString(CultureInfo.InvariantCulture), piece ?? "");
                }

                return string.IsNullOrEmpty(visible) ? null : visible;
            }
            catch
            {
                return null;
            }
        }

        private XElement FindNumPr(XElement paragraph)
        {
            XElement pPr = paragraph.Element(W + "pPr");
            if (pPr != null)
            {
                XElement direct = pPr.Element(W + "numPr");
                if (direct != null)
                {
                    return direct;
                }

                string styleId = Val(pPr.Element(W + "pStyle"));
                int guard = 0;
                while (!string.IsNullOrEmpty(styleId) && guard++ < 16)
                {
                    XElement style;
                    if (!_styles.TryGetValue(styleId, out style))
                    {
                        break;
                    }

                    XElement fromStyle = style.Element(W + "pPr")?.Element(W + "numPr");
                    if (fromStyle != null)
                    {
                        return fromStyle;
                    }

                    styleId = Val(style.Element(W + "basedOn"));
                }
            }

            return null;
        }

        private static XElement FindLvl(XElement abs, XElement num, int ilvl)
        {
            foreach (XElement ov in num.Elements(W + "lvlOverride"))
            {
                int ovl = 0;
                int.TryParse((string)ov.Attribute(W + "ilvl") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out ovl);
                if (ovl == ilvl && ov.Element(W + "lvl") != null)
                {
                    return ov.Element(W + "lvl");
                }
            }

            foreach (XElement lvl in abs.Elements(W + "lvl"))
            {
                int lv = 0;
                int.TryParse((string)lvl.Attribute(W + "ilvl") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out lv);
                if (lv == ilvl)
                {
                    return lvl;
                }
            }

            return abs.Element(W + "lvl");
        }

        private int StartOf(XElement num, XElement lvl, int ilvl)
        {
            foreach (XElement ov in num.Elements(W + "lvlOverride"))
            {
                int ovl = 0;
                int.TryParse((string)ov.Attribute(W + "ilvl") ?? "0", NumberStyles.Integer, CultureInfo.InvariantCulture, out ovl);
                if (ovl == ilvl)
                {
                    string so = Val(ov.Element(W + "startOverride"));
                    int n;
                    if (int.TryParse(so, NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                    {
                        return n;
                    }
                }
            }

            int start;
            if (int.TryParse(Val(lvl.Element(W + "start")) ?? "1", NumberStyles.Integer, CultureInfo.InvariantCulture, out start))
            {
                return start;
            }

            return 1;
        }

        private int Advance(string numId, int ilvl, int start)
        {
            int[] counters;
            if (!_counters.TryGetValue(numId, out counters))
            {
                counters = new int[9];
                for (int i = 0; i < counters.Length; i++)
                {
                    counters[i] = int.MinValue;
                }

                _counters[numId] = counters;
            }

            if (ilvl >= counters.Length)
            {
                ilvl = counters.Length - 1;
            }

            if (counters[ilvl] == int.MinValue)
            {
                counters[ilvl] = start;
            }
            else
            {
                counters[ilvl]++;
            }

            for (int i = ilvl + 1; i < counters.Length; i++)
            {
                counters[i] = int.MinValue;
            }

            return counters[ilvl];
        }

        private int Current(string numId, int ilvl)
        {
            int[] counters;
            if (!_counters.TryGetValue(numId, out counters) || ilvl < 0 || ilvl >= counters.Length)
            {
                return 1;
            }

            return counters[ilvl] == int.MinValue ? 1 : counters[ilvl];
        }

        private static string FormatNumber(string fmt, int n, string lvlText)
        {
            if (string.Equals(fmt, "bullet", StringComparison.OrdinalIgnoreCase))
            {
                return "•";
            }

            if (string.Equals(fmt, "decimal", StringComparison.OrdinalIgnoreCase)
                || string.Equals(fmt, "decimalZero", StringComparison.OrdinalIgnoreCase))
            {
                return n.ToString(CultureInfo.InvariantCulture);
            }

            if (string.Equals(fmt, "upperLetter", StringComparison.OrdinalIgnoreCase))
            {
                return ToLetters(n, 'A');
            }

            if (string.Equals(fmt, "lowerLetter", StringComparison.OrdinalIgnoreCase))
            {
                return ToLetters(n, 'a');
            }

            if (string.Equals(fmt, "upperRoman", StringComparison.OrdinalIgnoreCase))
            {
                return ToRoman(n).ToUpperInvariant();
            }

            if (string.Equals(fmt, "lowerRoman", StringComparison.OrdinalIgnoreCase))
            {
                return ToRoman(n).ToLowerInvariant();
            }

            return n.ToString(CultureInfo.InvariantCulture);
        }

        private static string ToLetters(int n, char a)
        {
            if (n <= 0)
            {
                return a.ToString();
            }

            var chars = new Stack<char>();
            int x = n;
            while (x > 0)
            {
                x--;
                chars.Push((char)(a + (x % 26)));
                x /= 26;
            }

            return new string(chars.ToArray());
        }

        private static string ToRoman(int n)
        {
            if (n <= 0)
            {
                return "I";
            }

            var map = new[]
            {
                Tuple.Create(1000, "M"), Tuple.Create(900, "CM"), Tuple.Create(500, "D"),
                Tuple.Create(400, "CD"), Tuple.Create(100, "C"), Tuple.Create(90, "XC"),
                Tuple.Create(50, "L"), Tuple.Create(40, "XL"), Tuple.Create(10, "X"),
                Tuple.Create(9, "IX"), Tuple.Create(5, "V"), Tuple.Create(4, "IV"),
                Tuple.Create(1, "I")
            };
            var sb = new System.Text.StringBuilder();
            int left = n;
            foreach (var pair in map)
            {
                while (left >= pair.Item1)
                {
                    sb.Append(pair.Item2);
                    left -= pair.Item1;
                }
            }

            return sb.ToString();
        }

        private static string Val(XElement el)
        {
            if (el == null)
            {
                return null;
            }

            return (string)el.Attribute(W + "val");
        }

        private static XElement FindPartRoot(XDocument pkg, string fileName, XName root)
        {
            foreach (XElement part in pkg.Descendants(Pkg + "part"))
            {
                string name = (string)part.Attribute(Pkg + "name");
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                if (name.EndsWith("/" + fileName, StringComparison.OrdinalIgnoreCase)
                    || name.EndsWith(fileName, StringComparison.OrdinalIgnoreCase))
                {
                    XElement data = part.Element(Pkg + "xmlData");
                    if (data != null)
                    {
                        return data.Element(root) ?? data.Element(root.LocalName);
                    }
                }
            }

            return pkg.Root != null && pkg.Root.Name == root ? pkg.Root : null;
        }
    }
}
