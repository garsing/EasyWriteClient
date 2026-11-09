using System;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1.PresentationHost
{
    /// <summary>
    /// 约定 HTML 的 SVG：校验后的原文进会话 ppt_svg，形状上只记模式和相对路径。
    /// </summary>
    internal static class PptHtmlSvgStore
    {
        private const string ModeTag = "ew-svg-mode";
        private const string FileTag = "ew-svg-file";
        private const string SrcTag = "ew-svg-src";

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
            object shape,
            string slideId,
            PptHtmlApplyNode node,
            bool linked,
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

            if (!TryPersist(slideId, comId, node.SvgMarkup, out string relative, out string full, out error))
            {
                return false;
            }

            if (linked && !TryRetargetLink(shape, full))
            {
                error = "插入 svg 后不能改到工作区文件";
                return false;
            }

            string mode = string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase)
                ? "inline"
                : "file";
            return TryStamp(shape, mode, relative, node.DataSrc, out error);
        }

        /// <summary>
        /// 已有链接时覆盖 ppt_svg 并改链接，ShapeId 不变。
        /// 没有链接时 recreate=true，调用方删掉再插。
        /// </summary>
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

            if (!PptShapeId.TryParseShape(node.ShapeId, out string slideId, out int comId))
            {
                error = "ShapeId 无法解析";
                return false;
            }

            if (!TryReadLinkPath(shape, out _))
            {
                recreate = true;
                return true;
            }

            string full = FullPath(slideId, comId);
            string previous = null;
            if (File.Exists(full))
            {
                previous = File.ReadAllText(full, Encoding.UTF8);
            }

            if (!TryPersist(slideId, comId, node.SvgMarkup, out string relative, out full, out error))
            {
                return false;
            }

            if (!TryRetargetLink(shape, full))
            {
                Restore(full, previous);
                recreate = true;
                return true;
            }

            string mode = string.Equals(node.ShapeType, "svg", StringComparison.OrdinalIgnoreCase)
                ? "inline"
                : "file";
            if (!TryStamp(shape, mode, relative, node.DataSrc, out error))
            {
                Restore(full, previous);
                return false;
            }

            return true;
        }

        public static bool TryConsumeRead(
            object shape,
            ref string typeName,
            out string mode,
            out string markup,
            out string src,
            out string error)
        {
            mode = null;
            markup = null;
            src = null;
            error = null;
            if (!TryReadMarks(shape, out mode, out string relative, out src))
            {
                mode = null;
                src = null;
                return true;
            }

            if (string.Equals(mode, "inline", StringComparison.OrdinalIgnoreCase))
            {
                if (!TryLoad(relative, out markup, out error))
                {
                    error = "svg 原文丢失: " + (error ?? "");
                    return false;
                }

                typeName = "svg";
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
            try
            {
                el = XElement.Parse(node.SvgMarkup);
            }
            catch (Exception)
            {
                el = new XElement("svg");
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

        private static bool TryPersist(
            string slideId,
            int comId,
            string markup,
            out string relative,
            out string full,
            out string error)
        {
            relative = null;
            full = null;
            error = null;
            try
            {
                string dir = Path.Combine(WorkspacePathResolver.GetSessionDirectory(), "ppt_svg");
                Directory.CreateDirectory(dir);
                string name = SafeId(slideId) + "_" + comId.ToString(CultureInfo.InvariantCulture) + ".svg";
                full = Path.Combine(dir, name);
                File.WriteAllText(full, markup, new UTF8Encoding(false));
                relative = "ppt_svg/" + name;
                return true;
            }
            catch (Exception ex)
            {
                error = "保存 svg 原文失败: " + ex.Message;
                return false;
            }
        }

        private static bool TryLoad(string relative, out string markup, out string error)
        {
            markup = null;
            error = null;
            if (string.IsNullOrWhiteSpace(relative) || relative.Contains(".."))
            {
                error = "svg 路径无效";
                return false;
            }

            try
            {
                string session = WorkspacePathResolver.GetSessionDirectory();
                string root = Path.GetFullPath(Path.Combine(session, "ppt_svg"));
                string full = Path.GetFullPath(Path.Combine(
                    session,
                    relative.Replace('/', Path.DirectorySeparatorChar)));
                if (!full.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(full, root, StringComparison.OrdinalIgnoreCase))
                {
                    error = "svg 路径无效";
                    return false;
                }

                if (!File.Exists(full))
                {
                    error = full;
                    return false;
                }

                markup = File.ReadAllText(full, Encoding.UTF8);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static string FullPath(string slideId, int comId)
        {
            string name = SafeId(slideId) + "_" + comId.ToString(CultureInfo.InvariantCulture) + ".svg";
            return Path.Combine(WorkspacePathResolver.GetSessionDirectory(), "ppt_svg", name);
        }

        private static void Restore(string full, string previous)
        {
            try
            {
                if (previous == null)
                {
                    if (File.Exists(full))
                    {
                        File.Delete(full);
                    }

                    return;
                }

                File.WriteAllText(full, previous, new UTF8Encoding(false));
            }
            catch (Exception)
            {
            }
        }

        private static string SafeId(string slideId)
        {
            var sb = new StringBuilder();
            if (slideId != null)
            {
                for (int i = 0; i < slideId.Length; i++)
                {
                    char c = slideId[i];
                    if (char.IsLetterOrDigit(c) || c == '-' || c == '_')
                    {
                        sb.Append(c);
                    }
                }
            }

            return sb.Length == 0 ? "slide" : sb.ToString();
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
                return !string.IsNullOrWhiteSpace(relative);
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

        private static bool TryRetargetLink(object shape, string fullPath)
        {
            try
            {
                object link = shape.GetType().InvokeMember(
                    "LinkFormat",
                    BindingFlags.GetProperty,
                    null,
                    shape,
                    null);
                if (link == null)
                {
                    return false;
                }

                link.GetType().InvokeMember(
                    "SourceFullName",
                    BindingFlags.SetProperty,
                    null,
                    link,
                    new object[] { fullPath });
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static bool TryReadLinkPath(object shape, out string path)
        {
            path = null;
            try
            {
                object link = shape.GetType().InvokeMember(
                    "LinkFormat",
                    BindingFlags.GetProperty,
                    null,
                    shape,
                    null);
                if (link == null)
                {
                    return false;
                }

                object src = link.GetType().InvokeMember(
                    "SourceFullName",
                    BindingFlags.GetProperty,
                    null,
                    link,
                    null);
                path = Convert.ToString(src);
                return !string.IsNullOrWhiteSpace(path);
            }
            catch (Exception)
            {
                return false;
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
