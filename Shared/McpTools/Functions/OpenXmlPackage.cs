using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>已打开文档的 Flat OPC 快照。读失败直接抛，不回退 COM。</summary>
    public sealed class OpenXmlPackage
    {
        public static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public static readonly XNamespace R =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        public static readonly XNamespace Pkg =
            "http://schemas.microsoft.com/office/2006/xmlPackage";

        public static readonly XNamespace Rel =
            "http://schemas.openxmlformats.org/package/2006/relationships";

        public XDocument Xml { get; private set; }

        public XElement Body { get; private set; }

        public static OpenXmlPackage Load(Word.Document document)
        {
            if (document == null)
            {
                throw new InvalidOperationException("文档为空，无法导出 WordOpenXML");
            }

            string raw;
            try
            {
                raw = document.Content.WordOpenXML;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("导出 WordOpenXML 失败: " + ex.Message, ex);
            }

            if (string.IsNullOrEmpty(raw))
            {
                throw new InvalidOperationException("WordOpenXML 为空");
            }

            XDocument xml;
            try
            {
                xml = XDocument.Parse(raw);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("解析 WordOpenXML 失败: " + ex.Message, ex);
            }

            XElement body = xml.Descendants(W + "body").FirstOrDefault();
            if (body == null)
            {
                throw new InvalidOperationException("WordOpenXML 中没有 w:body");
            }

            return new OpenXmlPackage
            {
                Xml = xml,
                Body = body
            };
        }

        public XElement FindPartRoot(string localNameEndsWith, XName rootName)
        {
            foreach (XElement part in Xml.Descendants(Pkg + "part"))
            {
                string name = (string)part.Attribute(Pkg + "name") ?? "";
                if (name.IndexOf(localNameEndsWith, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    XElement root = part.Descendants(rootName).FirstOrDefault();
                    if (root != null)
                    {
                        return root;
                    }
                }
            }

            return Xml.Descendants(rootName).FirstOrDefault();
        }

        public Dictionary<string, string> LoadDocumentRels()
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            XElement rels = FindPartRoot("document.xml.rels", Rel + "Relationships");
            if (rels == null)
            {
                return map;
            }

            foreach (XElement rel in rels.Elements(Rel + "Relationship"))
            {
                string id = (string)rel.Attribute("Id") ?? "";
                string target = (string)rel.Attribute("Target") ?? "";
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                {
                    map[id] = target.Replace('\\', '/');
                }
            }

            return map;
        }

        public XElement FindWordPartByTarget(string target)
        {
            if (string.IsNullOrEmpty(target))
            {
                return null;
            }

            string file = target;
            int slash = file.LastIndexOf('/');
            if (slash >= 0)
            {
                file = file.Substring(slash + 1);
            }

            foreach (XElement part in Xml.Descendants(Pkg + "part"))
            {
                string name = (string)part.Attribute(Pkg + "name") ?? "";
                if (name.EndsWith("/" + file, StringComparison.OrdinalIgnoreCase)
                    || name.EndsWith(file, StringComparison.OrdinalIgnoreCase))
                {
                    return part.Elements().FirstOrDefault()
                        ?? part.Descendants().FirstOrDefault();
                }
            }

            return null;
        }
    }
}
