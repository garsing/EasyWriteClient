using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 约定 HTML 的 SVG。内嵌原文进演示文稿的自定义 XML 部件，形状上只记模式。
    /// </summary>
    internal static class PptHtmlSvgStore
    {
        private const string ModeTag = "ew-svg-mode";
        private const string FileTag = "ew-svg-file";
        private const string SrcTag = "ew-svg-src";
        private const string PartNs = "urn:easywrite:ppt-svg:v1";
        private const string MissingText = "svg 原文丢失";
        private static readonly XNamespace Ew = PartNs;

        public static bool IsSvgNode(PptHtmlApplyNode node)
        {
            if (node == null)
            {
                return false;
            }

            if (string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return string.Equals(node.SvgMode, "inline", StringComparison.OrdinalIgnoreCase)
                || string.Equals(node.SvgMode, "file", StringComparison.OrdinalIgnoreCase);
        }

        public static bool IsSvgPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            string trimmed = path.Trim();
            int query = trimmed.IndexOf('?');
            if (query >= 0)
            {
                trimmed = trimmed.Substring(0, query);
            }

            return trimmed.EndsWith(".svg", StringComparison.OrdinalIgnoreCase);
        }

        public static bool TryValidateMarkup(string markup, out string normalized, out string error)
        {
            normalized = null;
            SvgSecurityValidator.ValidationResult result = SvgSecurityValidator.Validate(markup);
            if (result == null || !result.Success)
            {
                error = result == null ? "svg 未通过校验" : (result.Error ?? "svg 未通过校验");
                return false;
            }

            normalized = result.NormalizedSvg;
            error = null;
            return true;
        }

        /// <summary>已解析出的本地 .svg。不是 svg 则返回 null。</summary>
        public static string TryAcceptFile(PptHtmlApplyNode node)
        {
            if (node == null)
            {
                return null;
            }

            if (string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            if (!IsSvgPath(node.DataSrc) && !IsSvgPath(node.ResolvedLocalPath))
            {
                return null;
            }

            if (string.IsNullOrEmpty(node.ResolvedLocalPath) || !File.Exists(node.ResolvedLocalPath))
            {
                return "svg 文件不存在: " + (node.DataSrc ?? "");
            }

            if (!TryReadLocalAndValidate(node.ResolvedLocalPath, out string normalized, out string error))
            {
                return error;
            }

            node.SvgMarkup = normalized;
            node.SvgMode = "file";
            return null;
        }

        /// <summary>
        /// WPS 的 AddPicture 拒绝这些元素。插入失败时点名，没有这些元素则返回 null。
        /// </summary>
        public static string DescribeWppUnsupported(string markup)
        {
            if (string.IsNullOrWhiteSpace(markup))
            {
                return null;
            }

            XElement root;
            try
            {
                root = XElement.Parse(markup);
            }
            catch (Exception)
            {
                return null;
            }

            var found = new List<string>();
            foreach (XElement el in root.Descendants())
            {
                string local = el.Name.LocalName;
                if (!string.Equals(local, "g", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(local, "text", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string tag = "<" + local.ToLowerInvariant() + ">";
                if (!found.Contains(tag))
                {
                    found.Add(tag);
                }
            }

            if (found.Count == 0)
            {
                return null;
            }

            return "WPS 不支持 " + string.Join("、", found);
        }

        public static bool TryPrepareInsertPath(
            PptHtmlApplyNode node,
            out string path,
            out bool deleteAfter,
            out string error)
        {
            path = null;
            deleteAfter = false;
            error = null;
            if (node == null || string.IsNullOrEmpty(node.SvgMarkup))
            {
                error = "缺少 svg 原文";
                return false;
            }

            if (string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(node.SvgMode, "inline", StringComparison.OrdinalIgnoreCase))
            {
                path = Path.Combine(Path.GetTempPath(), "ew-svg-" + Guid.NewGuid().ToString("N") + ".svg");
                File.WriteAllText(path, node.SvgMarkup, new UTF8Encoding(false));
                deleteAfter = true;
                return true;
            }

            if (string.IsNullOrEmpty(node.ResolvedLocalPath) || !File.Exists(node.ResolvedLocalPath))
            {
                error = "svg 文件不存在: " + (node.DataSrc ?? "");
                return false;
            }

            path = Path.GetFullPath(node.ResolvedLocalPath);
            return true;
        }

        public static void DeleteQuiet(string path, bool delete)
        {
            if (!delete || string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch (Exception)
            {
            }
        }

        public static bool TryFinishInserted(
            object presentation,
            object shape,
            string slideId,
            PptHtmlApplyNode node,
            out string error)
        {
            error = null;
            int comId = ReadId(shape);
            if (comId <= 0)
            {
                error = "插入 svg 后没有形状编号";
                return false;
            }

            if (node == null || string.IsNullOrEmpty(node.SvgMarkup))
            {
                error = "缺少 svg 原文";
                return false;
            }

            bool inline = string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase)
                || string.Equals(node.SvgMode, "inline", StringComparison.OrdinalIgnoreCase);
            if (inline)
            {
                if (!TrySaveInline(presentation, slideId, comId, node.SvgMarkup, out error))
                {
                    return false;
                }

                return TryStamp(shape, "inline", "", "", out error);
            }

            return TryStamp(shape, "file", "", node.DataSrc, out error);
        }

        /// <summary>内嵌没有外部文件可改，调用方删掉再插。旧编号的部件条目由调用方删。</summary>
        public static bool TryReplaceSource(
            object shape,
            PptHtmlApplyNode node,
            out bool recreate,
            out string error)
        {
            recreate = false;
            error = null;
            if (node == null || string.IsNullOrEmpty(node.SvgMarkup))
            {
                error = "缺少 svg 原文";
                return false;
            }

            if (!PptShapeId.TryParseShape(node.ShapeId, out _, out _))
            {
                error = "ShapeId 无法解析";
                return false;
            }

            recreate = true;
            return true;
        }

        public static void ForgetInline(object presentation, string slideId, int comId)
        {
            try
            {
                TryMutate(presentation, root =>
                {
                    XElement item = FindItem(root, slideId, comId);
                    if (item != null)
                    {
                        item.Remove();
                    }
                }, out _);
            }
            catch (Exception)
            {
            }
        }

        public static bool TryConsumeRead(
            object shape,
            string slideId,
            int comId,
            ref string typeName,
            out string mode,
            out string markup,
            out string src,
            out bool sourceMissing,
            out string error)
        {
            mode = null;
            markup = null;
            src = null;
            sourceMissing = false;
            error = null;
            if (!TryReadMarks(shape, out mode, out _, out src))
            {
                mode = null;
                src = null;
                return true;
            }

            if (string.Equals(mode, "inline", StringComparison.OrdinalIgnoreCase))
            {
                typeName = "svg";
                if (!TryPresentation(shape, out object presentation, out error))
                {
                    return false;
                }

                if (!TryReadInline(presentation, slideId, comId, out markup, out bool missing, out error))
                {
                    return false;
                }

                if (missing)
                {
                    sourceMissing = true;
                    markup = null;
                    error = null;
                }

                return true;
            }

            if (string.Equals(mode, "file", StringComparison.OrdinalIgnoreCase))
            {
                typeName = "picture";
                return true;
            }

            mode = null;
            src = null;
            return true;
        }

        public static string FormatInlineElement(PptHtmlShapeNode node, string indent)
        {
            XElement el;
            if (node != null && node.SvgSourceMissing)
            {
                el = new XElement("svg");
                el.SetAttributeValue("data-error", MissingText);
            }
            else
            {
                try
                {
                    el = XElement.Parse(node.SvgMarkup);
                }
                catch (Exception)
                {
                    el = new XElement("svg");
                }

                el.SetAttributeValue("data-error", null);
            }

            if (!string.IsNullOrEmpty(node.ShapeId))
            {
                el.SetAttributeValue("ShapeId", node.ShapeId);
            }

            el.SetAttributeValue("data-shape-type", "svg");
            if (!string.IsNullOrEmpty(node.Style))
            {
                el.SetAttributeValue("style", node.Style);
            }

            return (indent ?? "") + el.ToString(SaveOptions.DisableFormatting) + Environment.NewLine;
        }

        private static bool TryReadLocalAndValidate(string localPath, out string normalized, out string error)
        {
            normalized = null;
            try
            {
                string text = File.ReadAllText(localPath, Encoding.UTF8);
                return TryValidateMarkup(text, out normalized, out error);
            }
            catch (Exception ex)
            {
                error = "读取 svg 失败: " + ex.Message;
                return false;
            }
        }

        private static bool TrySaveInline(
            object presentation,
            string slideId,
            int comId,
            string markup,
            out string error)
        {
            string payload = Convert.ToBase64String(Encoding.UTF8.GetBytes(markup ?? ""));
            return TryMutate(presentation, root =>
            {
                XElement item = FindItem(root, slideId, comId);
                if (item == null)
                {
                    item = new XElement(Ew + "item");
                    item.SetAttributeValue("slide", slideId ?? "");
                    item.SetAttributeValue("shape", comId.ToString(CultureInfo.InvariantCulture));
                    root.Add(item);
                }

                item.Value = payload;
            }, out error);
        }

        private static bool TryReadInline(
            object presentation,
            string slideId,
            int comId,
            out string markup,
            out bool missing,
            out string error)
        {
            markup = null;
            missing = false;
            error = null;
            if (!TryLoadRoot(presentation, out XElement root, out bool found, out error))
            {
                return false;
            }

            if (!found)
            {
                missing = true;
                return true;
            }

            XElement item = FindItem(root, slideId, comId);
            if (item == null || string.IsNullOrWhiteSpace(item.Value))
            {
                missing = true;
                return true;
            }

            try
            {
                byte[] raw = Convert.FromBase64String(item.Value.Trim());
                markup = Encoding.UTF8.GetString(raw);
                if (string.IsNullOrWhiteSpace(markup))
                {
                    missing = true;
                    markup = null;
                }

                return true;
            }
            catch (Exception)
            {
                missing = true;
                markup = null;
                return true;
            }
        }

        private static XElement FindItem(XElement root, string slideId, int comId)
        {
            if (root == null)
            {
                return null;
            }

            string shape = comId.ToString(CultureInfo.InvariantCulture);
            foreach (XElement item in root.Elements(Ew + "item"))
            {
                XAttribute slide = item.Attribute("slide");
                XAttribute id = item.Attribute("shape");
                if (slide != null
                    && id != null
                    && string.Equals(slide.Value, slideId ?? "", StringComparison.Ordinal)
                    && string.Equals(id.Value, shape, StringComparison.Ordinal))
                {
                    return item;
                }
            }

            return null;
        }

        private static bool TryMutate(object presentation, Action<XElement> mutate, out string error)
        {
            error = null;
            try
            {
                if (!TryFindPart(presentation, out object part, out XElement root, out error))
                {
                    return false;
                }

                if (part == null)
                {
                    root = new XElement(Ew + "svgSources");
                }

                mutate(root);
                string xml = root.ToString(SaveOptions.DisableFormatting);
                object parts = Member(presentation, "CustomXMLParts");
                if (part != null)
                {
                    try
                    {
                        Member(part, "Delete");
                    }
                    catch (Exception)
                    {
                    }
                }

                Member(parts, "Add", xml);
                return true;
            }
            catch (Exception ex)
            {
                error = "保存 svg 原文失败: " + Unwrap(ex).Message;
                return false;
            }
        }

        private static bool TryLoadRoot(
            object presentation,
            out XElement root,
            out bool found,
            out string error)
        {
            if (!TryFindPart(presentation, out _, out root, out error))
            {
                found = false;
                return false;
            }

            found = root != null;
            return true;
        }

        private static bool TryFindPart(
            object presentation,
            out object part,
            out XElement root,
            out string error)
        {
            part = null;
            root = null;
            error = null;
            try
            {
                object parts = Member(presentation, "CustomXMLParts");
                object selected = null;
                try
                {
                    selected = Member(parts, "SelectByNamespace", PartNs);
                }
                catch (Exception)
                {
                    selected = null;
                }

                if (selected != null)
                {
                    int count = Convert.ToInt32(Member(selected, "Count"), CultureInfo.InvariantCulture);
                    if (count >= 1)
                    {
                        part = Member(selected, "Item", 1);
                        root = XElement.Parse(Convert.ToString(Member(part, "XML")) ?? "");
                        return true;
                    }
                }

                int all = Convert.ToInt32(Member(parts, "Count"), CultureInfo.InvariantCulture);
                for (int i = 1; i <= all; i++)
                {
                    object candidate = Member(parts, "Item", i);
                    string xml = Convert.ToString(Member(candidate, "XML")) ?? "";
                    if (xml.IndexOf(PartNs, StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    part = candidate;
                    root = XElement.Parse(xml);
                    return true;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "读取 svg 原文失败: " + Unwrap(ex).Message;
                return false;
            }
        }

        private static Exception Unwrap(Exception ex)
        {
            if (ex is TargetInvocationException tie && tie.InnerException != null)
            {
                return tie.InnerException;
            }

            return ex;
        }

        private static bool TryPresentation(object shape, out object presentation, out string error)
        {
            presentation = null;
            error = null;
            object current = shape;
            for (int i = 0; i < 8 && current != null; i++)
            {
                object parent;
                try
                {
                    parent = Member(current, "Parent");
                }
                catch (Exception)
                {
                    parent = null;
                }

                if (parent == null)
                {
                    break;
                }

                try
                {
                    object parts = Member(parent, "CustomXMLParts");
                    if (parts != null)
                    {
                        presentation = parent;
                        return true;
                    }
                }
                catch (Exception)
                {
                }

                current = parent;
            }

            error = "找不到演示文稿";
            return false;
        }

        private static object Member(object target, string name, params object[] args)
        {
            if (args == null || args.Length == 0)
            {
                try
                {
                    return target.GetType().InvokeMember(
                        name,
                        BindingFlags.GetProperty | BindingFlags.InvokeMethod,
                        null,
                        target,
                        null);
                }
                catch (Exception)
                {
                    return target.GetType().InvokeMember(
                        name,
                        BindingFlags.InvokeMethod,
                        null,
                        target,
                        null);
                }
            }

            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod | BindingFlags.GetProperty,
                null,
                target,
                args);
        }

        private static bool TryStamp(
            object shape,
            string mode,
            string relative,
            string dataSrc,
            out string error)
        {
            error = null;
            try
            {
                object tags = shape.GetType().InvokeMember(
                    "Tags",
                    BindingFlags.GetProperty,
                    null,
                    shape,
                    null);
                SetTag(tags, ModeTag, mode ?? "");
                SetTag(tags, FileTag, relative ?? "");
                SetTag(tags, SrcTag, string.Equals(mode, "file", StringComparison.OrdinalIgnoreCase)
                    ? (dataSrc ?? "")
                    : "");
                return true;
            }
            catch (Exception ex)
            {
                error = "记下 svg 标记失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryReadMarks(object shape, out string mode, out string relative, out string src)
        {
            mode = ReadTag(shape, ModeTag);
            relative = ReadTag(shape, FileTag);
            src = ReadTag(shape, SrcTag);
            if (string.Equals(mode, "inline", StringComparison.OrdinalIgnoreCase)
                || string.Equals(mode, "file", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            mode = null;
            relative = null;
            src = null;
            return false;
        }

        private static void SetTag(object tags, string name, string value)
        {
            try
            {
                tags.GetType().InvokeMember(
                    "Delete",
                    BindingFlags.InvokeMethod,
                    null,
                    tags,
                    new object[] { name });
            }
            catch (Exception)
            {
            }

            tags.GetType().InvokeMember(
                "Add",
                BindingFlags.InvokeMethod,
                null,
                tags,
                new object[] { name, value ?? "" });
        }

        private static string ReadTag(object shape, string name)
        {
            try
            {
                object tags = shape.GetType().InvokeMember(
                    "Tags",
                    BindingFlags.GetProperty,
                    null,
                    shape,
                    null);
                if (tags == null)
                {
                    return "";
                }

                object value;
                try
                {
                    value = tags.GetType().InvokeMember(
                        "Item",
                        BindingFlags.InvokeMethod | BindingFlags.GetProperty,
                        null,
                        tags,
                        new object[] { name });
                }
                catch (Exception)
                {
                    value = tags.GetType().InvokeMember(
                        "",
                        BindingFlags.GetProperty,
                        null,
                        tags,
                        new object[] { name });
                }

                return value == null ? "" : (Convert.ToString(value) ?? "");
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static int ReadId(object shape)
        {
            try
            {
                object id = shape.GetType().InvokeMember(
                    "Id",
                    BindingFlags.GetProperty,
                    null,
                    shape,
                    null);
                return Convert.ToInt32(id, CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                return 0;
            }
        }
    }
}
