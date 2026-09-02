using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Newtonsoft.Json.Linq;

namespace WordAddIn1.PresentationHost
{
    internal static class PptHtmlExtendsExpander
    {
        private static readonly HashSet<string> ForbiddenTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "smartart", "unknown", "freeform"
        };

        private static readonly Regex SlotRe = new Regex(
            @"^[a-z][a-z0-9_]{0,31}$",
            RegexOptions.CultureInvariant);

        public static bool TryExpand(string html, out string expanded, out string error)
        {
            expanded = html;
            error = null;
            if (string.IsNullOrWhiteSpace(html) || html.IndexOf("data-extends", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return true;
            }

            int sectionOpens = Regex.Matches(html, @"<section\b", RegexOptions.IgnoreCase).Count;
            if (sectionOpens == 0)
            {
                error = "html 中未找到 <section>";
                return false;
            }

            if (sectionOpens > 1)
            {
                error = "html 中只能有一个 <section>";
                return false;
            }

            Match m = Regex.Match(
                html,
                @"<section\b[^>]*>.*?</section>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!m.Success)
            {
                error = "无法解析 <section>…</section>";
                return false;
            }

            XElement section;
            try
            {
                section = XElement.Parse(m.Value, LoadOptions.PreserveWhitespace);
            }
            catch (Exception ex)
            {
                error = "section HTML 无法解析为 XML: " + ex.Message;
                return false;
            }

            foreach (XElement nested in section.Descendants().Where(el => !section.Elements().Contains(el)))
            {
                if (!string.IsNullOrEmpty(GetAttr(nested, "data-extends")))
                {
                    error = "data-extends 只能写在 section 直接子级";
                    return false;
                }
            }

            var replacements = new List<KeyValuePair<XElement, List<XElement>>>();
            foreach (XElement child in section.Elements().ToList())
            {
                string extends = GetAttr(child, "data-extends");
                if (string.IsNullOrEmpty(extends))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(GetAttr(child, "ShapeId")))
                {
                    error = "data-extends 节点不能带 ShapeId";
                    return false;
                }

                if (!TryExpandOne(child, extends, out List<XElement> nodes, out error))
                {
                    return false;
                }

                replacements.Add(new KeyValuePair<XElement, List<XElement>>(child, nodes));
            }

            foreach (KeyValuePair<XElement, List<XElement>> one in replacements)
            {
                foreach (XElement node in one.Value)
                {
                    one.Key.AddBeforeSelf(node);
                }

                one.Key.Remove();
            }

            expanded = html.Substring(0, m.Index) + section.ToString(SaveOptions.DisableFormatting)
                + html.Substring(m.Index + m.Length);
            return true;
        }

        private static bool TryExpandOne(
            XElement instance,
            string extends,
            out List<XElement> nodes,
            out string error)
        {
            nodes = null;
            error = null;
            if (!PptLibStore.TryParseId(extends, out string prefix, out string full) || prefix != "usr")
            {
                error = "data-extends 只许 usr- 私有 id";
                return false;
            }

            if (!PptHtmlApplyParser.TryParseGeometry(
                    GetAttr(instance, "style"),
                    out double instL,
                    out double instT,
                    out double instW,
                    out double instH))
            {
                error = "data-extends 节点必须有幻灯片百分比几何";
                return false;
            }

            if (!PptLibStore.TryReadComponentHtml(full, out string fragment, out JObject meta, out error))
            {
                return false;
            }

            XElement root;
            try
            {
                root = XElement.Parse("<root>" + fragment + "</root>", LoadOptions.PreserveWhitespace);
            }
            catch (Exception ex)
            {
                error = "component.html 无法解析: " + ex.Message;
                return false;
            }

            if (root.DescendantsAndSelf().Any(el => !string.IsNullOrEmpty(GetAttr(el, "data-extends"))))
            {
                error = "组件内禁止再写 data-extends";
                return false;
            }

            var componentSlots = new HashSet<string>(StringComparer.Ordinal);
            foreach (XElement el in root.DescendantsAndSelf())
            {
                string slot = GetAttr(el, "data-slot");
                if (string.IsNullOrEmpty(slot))
                {
                    continue;
                }

                if (!SlotRe.IsMatch(slot))
                {
                    error = "组件槽名非法: " + slot;
                    return false;
                }

                componentSlots.Add(slot);
            }

            var instanceSlots = new Dictionary<string, XElement>(StringComparer.Ordinal);
            foreach (XElement el in instance.DescendantsAndSelf())
            {
                if (el == instance)
                {
                    continue;
                }

                string slot = GetAttr(el, "data-slot");
                if (string.IsNullOrEmpty(slot))
                {
                    continue;
                }

                if (!SlotRe.IsMatch(slot) || !componentSlots.Contains(slot))
                {
                    error = "未知槽名: " + slot;
                    return false;
                }

                instanceSlots[slot] = el;
            }

            int? instZ = TryParseInt(GetAttr(instance, "data-z"));
            nodes = new List<XElement>();
            foreach (XElement child in root.Elements())
            {
                if (!TryValidateComponentNode(child, out error))
                {
                    return false;
                }

                if (!PptHtmlApplyParser.TryParseGeometry(
                        GetAttr(child, "style"),
                        out double cL,
                        out double cT,
                        out double cW,
                        out double cH))
                {
                    error = "组件节点缺几何";
                    return false;
                }

                XElement clone = new XElement(child);
                string shapeType = GetAttr(clone, "data-shape-type");
                bool isGroup = string.Equals(shapeType, "group", StringComparison.OrdinalIgnoreCase);
                if (isGroup)
                {
                    clone.SetAttributeValue("style", PptConventionHtml.BuildStyle(instL, instT, instW, instH));
                }
                else
                {
                    clone.SetAttributeValue(
                        "style",
                        PptConventionHtml.BuildStyle(
                            instL + cL / 100.0 * instW,
                            instT + cT / 100.0 * instH,
                            cW / 100.0 * instW,
                            cH / 100.0 * instH));
                }

                StripShapeIds(clone);
                ApplySlotsRecursive(clone, instanceSlots);
                if (instZ.HasValue)
                {
                    int childZ = TryParseInt(GetAttr(clone, "data-z")) ?? 0;
                    clone.SetAttributeValue("data-z", (childZ + instZ.Value).ToString(CultureInfo.InvariantCulture));
                }

                nodes.Add(clone);
            }

            if (nodes.Count == 0)
            {
                error = "组件没有可展开节点: " + full;
                return false;
            }

            return true;
        }

        private static bool TryValidateComponentNode(XElement el, out string error)
        {
            error = null;
            if (el == null)
            {
                return true;
            }

            string shapeType = GetAttr(el, "data-shape-type");
            bool isGroup = string.Equals(shapeType, "group", StringComparison.OrdinalIgnoreCase);
            if (ForbiddenTypes.Contains(shapeType)
                || (!isGroup && !PptShapeTypeMap.IsCreatable(
                    PptShapeTypeMap.NormalizeForCreate(shapeType, el.Name.LocalName))))
            {
                error = "组件含不可建类型: " + (shapeType ?? "");
                return false;
            }

            if (isGroup)
            {
                foreach (XElement child in el.Elements())
                {
                    if (!TryValidateComponentNode(child, out error))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static void StripShapeIds(XElement el)
        {
            if (el == null)
            {
                return;
            }

            foreach (XAttribute attr in el.Attributes()
                .Where(a => string.Equals(a.Name.LocalName, "ShapeId", StringComparison.OrdinalIgnoreCase))
                .ToList())
            {
                attr.Remove();
            }

            foreach (XElement child in el.Elements())
            {
                StripShapeIds(child);
            }
        }

        private static void ApplySlotsRecursive(
            XElement el,
            Dictionary<string, XElement> instanceSlots)
        {
            if (el == null || instanceSlots == null)
            {
                return;
            }

            string slot = GetAttr(el, "data-slot");
            if (!string.IsNullOrEmpty(slot) && instanceSlots.TryGetValue(slot, out XElement overrideEl))
            {
                ApplySlotOverride(el, overrideEl);
            }

            foreach (XElement child in el.Elements())
            {
                ApplySlotsRecursive(child, instanceSlots);
            }
        }

        private static void ApplySlotOverride(XElement target, XElement source)
        {
            foreach (string attr in new[] { "data-font-color", "data-font-size", "data-font-bold", "data-valign", "data-align" })
            {
                string val = GetAttr(source, attr);
                if (!string.IsNullOrEmpty(val))
                {
                    target.SetAttributeValue(attr, val);
                }
            }

            string text = GetElementText(source);
            if (text != null)
            {
                target.RemoveNodes();
                target.Add(new XText(text));
            }
        }

        private static int? TryParseInt(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            if (int.TryParse(raw.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int n))
            {
                return n;
            }

            return null;
        }

        private static string GetAttr(XElement el, string name)
        {
            if (el == null)
            {
                return "";
            }

            XAttribute a = el.Attribute(name)
                ?? el.Attributes().FirstOrDefault(x =>
                    string.Equals(x.Name.LocalName, name, StringComparison.OrdinalIgnoreCase));
            return a == null ? "" : (a.Value ?? "").Trim();
        }

        private static string GetElementText(XElement el)
        {
            if (el == null)
            {
                return "";
            }

            return string.Concat(el.DescendantNodes().OfType<XText>().Select(t => t.Value)).TrimEnd();
        }
    }
}
