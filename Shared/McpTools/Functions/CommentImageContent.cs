using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 从批注或回复的 Range 抽出按原文顺序排列的字和图。不走正文抽图，不发 I_。
    /// </summary>
    internal static class CommentImageContent
    {
        internal const string NoteMissing = "图片未取出";
        internal const string NoteHost = "该宿主未带出批注图片";
        internal const string NoteOverCap = "未看，超过 9 张";
        internal const int MaxImages = 9;

        private static readonly XNamespace Pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
        private static readonly XNamespace RelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string RawKey = "_raw";

        internal static List<Dictionary<string, object>> Read(object range, string host)
        {
            string fallback = StripPictureControls(ReadRangeText(range));
            int? inline = TryCount(range, "InlineShapes");
            int? shapes = TryCount(range, "ShapeRange");
            bool bothZero = inline == 0 && shapes == 0;
            if (bothZero)
            {
                return TextOnly(fallback);
            }

            int known = (inline ?? 0) + (shapes ?? 0);
            string xml = ReadWordOpenXml(range);
            if (string.IsNullOrWhiteSpace(xml))
            {
                return known > 0 ? TextPlusMissing(fallback, host) : TextOnly(fallback);
            }

            try
            {
                List<Dictionary<string, object>> segments = ReadPackage(xml, host);
                bool sawImage = segments.Exists(seg =>
                    string.Equals(seg["type"] as string, "image", StringComparison.Ordinal));
                if (!sawImage && known > 0)
                {
                    return TextPlusMissing(fallback, host);
                }

                if (!sawImage)
                {
                    return TextOnly(fallback);
                }

                return segments;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CommentImage] XML: " + ex.Message);
                return known > 0 ? TextPlusMissing(fallback, host) : TextOnly(fallback);
            }
        }

        internal static void ApplySendBudget(List<Dictionary<string, object>> rows)
        {
            var pending = new List<Dictionary<string, object>>();
            if (rows == null)
            {
                return;
            }

            foreach (Dictionary<string, object> row in rows)
            {
                CollectRaw(row, pending);
                if (row != null && row.TryGetValue("replies", out object repliesRaw)
                    && repliesRaw is List<Dictionary<string, object>> replies)
                {
                    foreach (Dictionary<string, object> reply in replies)
                    {
                        CollectRaw(reply, pending);
                    }
                }
            }

            int sent = 0;
            foreach (Dictionary<string, object> segment in pending)
            {
                byte[] raw = segment[RawKey] as byte[];
                segment.Remove(RawKey);
                if (raw == null || raw.Length == 0)
                {
                    continue;
                }

                if (sent >= MaxImages)
                {
                    segment["note"] = NoteOverCap;
                    continue;
                }

                try
                {
                    ImageCaptureCompressor.CompressedImage image = ImageCaptureCompressor.CompressBytes(raw);
                    segment["image_base64"] = Convert.ToBase64String(image.Bytes ?? Array.Empty<byte>());
                    segment["format"] = image.Format ?? "jpeg";
                    segment["width_px"] = image.Width;
                    segment["height_px"] = image.Height;
                    sent++;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[CommentImage] 压缩: " + ex.Message);
                    segment["note"] = NoteMissing;
                }
            }
        }

        internal static string StripPictureControls(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            var sb = new StringBuilder(text.Length);
            foreach (char c in text)
            {
                if (c == '\u0001' || c == '\u0008' || c == '\uFFFC')
                {
                    continue;
                }

                sb.Append(c);
            }

            return sb.ToString();
        }

        private static void CollectRaw(Dictionary<string, object> owner, List<Dictionary<string, object>> pending)
        {
            if (owner == null || !owner.TryGetValue("content", out object raw)
                || !(raw is List<Dictionary<string, object>> content))
            {
                return;
            }

            foreach (Dictionary<string, object> segment in content)
            {
                if (segment != null && segment.ContainsKey(RawKey))
                {
                    pending.Add(segment);
                }
            }
        }

        private static List<Dictionary<string, object>> TextOnly(string text)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "text",
                    ["text"] = text ?? ""
                }
            };
        }

        private static List<Dictionary<string, object>> TextPlusMissing(string text, string host)
        {
            var rows = new List<Dictionary<string, object>>();
            if (!string.IsNullOrEmpty(text))
            {
                rows.Add(new Dictionary<string, object>
                {
                    ["type"] = "text",
                    ["text"] = text
                });
            }

            rows.Add(MissingImage(host));
            return rows;
        }

        private static Dictionary<string, object> MissingImage(string host)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "image",
                ["note"] = MissingNote(host)
            };
        }

        private static string MissingNote(string host)
        {
            return string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase)
                ? NoteHost
                : NoteMissing;
        }

        private static string ReadRangeText(object range)
        {
            if (range == null)
            {
                return "";
            }

            try
            {
                return ((dynamic)range).Text ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static string ReadWordOpenXml(object range)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                return ((dynamic)range).WordOpenXML as string;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[CommentImage] WordOpenXML: " + ex.Message);
                return null;
            }
        }

        private static int? TryCount(object range, string propertyName)
        {
            if (range == null)
            {
                return null;
            }

            try
            {
                dynamic bag = ((dynamic)range).GetType().InvokeMember(
                    propertyName,
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    range,
                    null);
                return (int)bag.Count;
            }
            catch (Exception)
            {
                try
                {
                    dynamic value = propertyName == "InlineShapes"
                        ? ((dynamic)range).InlineShapes.Count
                        : ((dynamic)range).ShapeRange.Count;
                    return (int)value;
                }
                catch (Exception)
                {
                    return null;
                }
            }
        }

        private static List<Dictionary<string, object>> ReadPackage(string xml, string host)
        {
            XDocument doc = XDocument.Parse(xml);
            XElement body = FindBody(doc);
            Dictionary<string, string> rels = IndexRels(doc);
            Dictionary<string, byte[]> binaries = IndexBinaries(doc);
            var segments = new List<Dictionary<string, object>>();
            var buffer = new StringBuilder();
            if (body != null)
            {
                Walk(body, buffer, segments, rels, binaries, host);
            }

            FlushText(buffer, segments);
            return segments;
        }

        private static XElement FindBody(XDocument doc)
        {
            if (doc == null || doc.Root == null)
            {
                return null;
            }

            if (doc.Root.Name == Pkg + "package")
            {
                foreach (XElement part in doc.Root.Elements(Pkg + "part"))
                {
                    string name = part.Attribute(Pkg + "name")?.Value ?? "";
                    if (!name.EndsWith("/word/document.xml", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    XElement xmlData = part.Element(Pkg + "xmlData");
                    XElement inner = xmlData == null ? null : xmlData.Elements().FirstOrDefault();
                    XElement body = inner == null ? null : inner.Element(W + "body");
                    if (body != null)
                    {
                        return body;
                    }
                }
            }

            return doc.Descendants(W + "body").FirstOrDefault();
        }

        private static Dictionary<string, string> IndexRels(XDocument doc)
        {
            var rels = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (doc?.Root == null || doc.Root.Name != Pkg + "package")
            {
                return rels;
            }

            foreach (XElement part in doc.Root.Elements(Pkg + "part"))
            {
                string name = part.Attribute(Pkg + "name")?.Value ?? "";
                if (!name.EndsWith("/word/_rels/document.xml.rels", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                XElement xmlData = part.Element(Pkg + "xmlData");
                if (xmlData == null)
                {
                    continue;
                }

                foreach (XElement rel in xmlData.Descendants(RelNs + "Relationship"))
                {
                    string id = rel.Attribute("Id")?.Value;
                    string target = rel.Attribute("Target")?.Value;
                    if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                    {
                        rels[id] = target;
                    }
                }
            }

            return rels;
        }

        private static Dictionary<string, byte[]> IndexBinaries(XDocument doc)
        {
            var binaries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
            if (doc?.Root == null || doc.Root.Name != Pkg + "package")
            {
                return binaries;
            }

            foreach (XElement part in doc.Root.Elements(Pkg + "part"))
            {
                string name = part.Attribute(Pkg + "name")?.Value ?? "";
                XElement binary = part.Element(Pkg + "binaryData");
                if (binary == null || string.IsNullOrWhiteSpace(binary.Value) || string.IsNullOrEmpty(name))
                {
                    continue;
                }

                try
                {
                    string b64 = System.Text.RegularExpressions.Regex.Replace(binary.Value, @"\s+", "");
                    binaries[name] = Convert.FromBase64String(b64);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[CommentImage] 部件解码失败: " + name + " " + ex.Message);
                }
            }

            return binaries;
        }

        private static void Walk(
            XElement element,
            StringBuilder buffer,
            List<Dictionary<string, object>> segments,
            Dictionary<string, string> rels,
            Dictionary<string, byte[]> binaries,
            string host)
        {
            if (element == null)
            {
                return;
            }

            if (element.Name.LocalName == "AlternateContent")
            {
                XElement choice = element.Elements().FirstOrDefault(child => child.Name.LocalName == "Choice");
                XElement fallback = element.Elements().FirstOrDefault(child => child.Name.LocalName == "Fallback");
                Walk(choice ?? fallback, buffer, segments, rels, binaries, host);
                return;
            }

            if (element.Name == W + "drawing" || element.Name == W + "pict")
            {
                FlushText(buffer, segments);
                segments.Add(ImageSegment(element, rels, binaries, host));
                return;
            }

            if (element.Name == W + "t" || element.Name == W + "delText")
            {
                buffer.Append(StripPictureControls(element.Value));
                return;
            }

            if (element.Name == W + "tab")
            {
                buffer.Append('\t');
                return;
            }

            if (element.Name == W + "br" || element.Name == W + "cr")
            {
                buffer.Append('\n');
                return;
            }

            foreach (XElement child in element.Elements())
            {
                Walk(child, buffer, segments, rels, binaries, host);
            }

            if (element.Name == W + "p")
            {
                buffer.Append('\r');
            }
        }

        private static void FlushText(StringBuilder buffer, List<Dictionary<string, object>> segments)
        {
            if (buffer.Length == 0)
            {
                return;
            }

            string text = buffer.ToString();
            buffer.Clear();
            if (text.Length == 0)
            {
                return;
            }

            segments.Add(new Dictionary<string, object>
            {
                ["type"] = "text",
                ["text"] = text
            });
        }

        private static Dictionary<string, object> ImageSegment(
            XElement drawing,
            Dictionary<string, string> rels,
            Dictionary<string, byte[]> binaries,
            string host)
        {
            byte[] bytes = ResolveBytes(drawing, rels, binaries);
            var segment = new Dictionary<string, object> { ["type"] = "image" };
            if (bytes == null || bytes.Length == 0)
            {
                segment["note"] = MissingNote(host);
                return segment;
            }

            segment[RawKey] = bytes;
            return segment;
        }

        private static byte[] ResolveBytes(
            XElement drawing,
            Dictionary<string, string> rels,
            Dictionary<string, byte[]> binaries)
        {
            foreach (XElement node in drawing.Descendants())
            {
                string relId = node.Attribute(R + "embed")?.Value
                    ?? node.Attribute(R + "link")?.Value
                    ?? node.Attribute(R + "id")?.Value;
                if (string.IsNullOrEmpty(relId))
                {
                    continue;
                }

                if (rels == null || !rels.TryGetValue(relId, out string target) || string.IsNullOrEmpty(target))
                {
                    continue;
                }

                byte[] bytes = LookupBinary(binaries, target);
                if (bytes != null && bytes.Length > 0)
                {
                    return bytes;
                }
            }

            return null;
        }

        private static byte[] LookupBinary(Dictionary<string, byte[]> binaries, string target)
        {
            if (binaries == null || binaries.Count == 0 || string.IsNullOrEmpty(target))
            {
                return null;
            }

            string cleaned = target.Replace('\\', '/');
            var candidates = new List<string>();
            if (cleaned.StartsWith("/", StringComparison.Ordinal))
            {
                candidates.Add(cleaned);
            }
            else if (cleaned.StartsWith("word/", StringComparison.OrdinalIgnoreCase))
            {
                candidates.Add("/" + cleaned);
            }
            else if (cleaned.StartsWith("media/", StringComparison.OrdinalIgnoreCase))
            {
                candidates.Add("/word/" + cleaned);
            }
            else
            {
                candidates.Add("/word/" + cleaned.TrimStart('/'));
                candidates.Add("/word/media/" + cleaned.TrimStart('/'));
            }

            foreach (string name in candidates)
            {
                if (binaries.TryGetValue(name, out byte[] bytes))
                {
                    return bytes;
                }
            }

            string file = cleaned;
            int slash = file.LastIndexOf('/');
            if (slash >= 0)
            {
                file = file.Substring(slash + 1);
            }

            foreach (KeyValuePair<string, byte[]> pair in binaries)
            {
                if (pair.Key.EndsWith("/" + file, StringComparison.OrdinalIgnoreCase))
                {
                    return pair.Value;
                }
            }

            return null;
        }
    }
}
