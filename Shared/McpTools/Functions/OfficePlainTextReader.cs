using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using PdfiumViewer;

namespace WordAddIn1
{
    /// <summary>
    /// 办公文档 / PDF 简便模式：按阅读顺序拉平纯文本，不写 table / S_ 标签。
    /// </summary>
    internal static class OfficePlainTextReader
    {
        private static readonly XNamespace W =
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private static readonly XNamespace A =
            "http://schemas.openxmlformats.org/drawingml/2006/main";
        private static readonly XNamespace Rels =
            "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace R =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace Main =
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        public static string Extract(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException("文件不存在", path);
            }

            string ext = FileReadClassifier.ExtensionOf(path);
            switch (ext)
            {
                case "docx":
                    return ExtractDocx(path);
                case "xlsx":
                case "xlsm":
                    return ExtractXlsx(path);
                case "pptx":
                    return ExtractPptx(path);
                case "pdf":
                    return ExtractPdf(path);
                case "doc":
                    return ExtractWordCom(path);
                case "xls":
                case "xlsb":
                    return ExtractExcelCom(path);
                case "ppt":
                    return ExtractPptCom(path);
                default:
                    throw new InvalidOperationException("不支持的办公格式: " + ext);
            }
        }

        private static string ExtractDocx(string path)
        {
            using (ZipArchive zip = ZipFile.OpenRead(path))
            {
                ZipArchiveEntry entry = zip.GetEntry("word/document.xml");
                if (entry == null)
                {
                    throw new InvalidOperationException("不是有效的 docx");
                }

                XDocument doc;
                using (Stream stream = entry.Open())
                {
                    doc = XDocument.Load(stream);
                }

                XElement body = doc.Root == null ? null : doc.Root.Element(W + "body");
                if (body == null)
                {
                    return "";
                }

                var sb = new StringBuilder();
                foreach (XElement child in body.Elements())
                {
                    if (child.Name == W + "p")
                    {
                        AppendWordParagraph(sb, child);
                        sb.Append('\n');
                    }
                    else if (child.Name == W + "tbl")
                    {
                        AppendWordTable(sb, child);
                    }
                }

                return sb.ToString().TrimEnd();
            }
        }

        private static void AppendWordParagraph(StringBuilder sb, XElement paragraph)
        {
            foreach (XElement node in paragraph.Descendants())
            {
                if (node.Ancestors(W + "del").Any())
                {
                    continue;
                }

                if (node.Name == W + "t")
                {
                    sb.Append(node.Value);
                }
                else if (node.Name == W + "tab")
                {
                    sb.Append('\t');
                }
                else if (node.Name == W + "br" || node.Name == W + "cr")
                {
                    sb.Append('\n');
                }
            }
        }

        private static void AppendWordTable(StringBuilder sb, XElement table)
        {
            foreach (XElement row in table.Elements(W + "tr"))
            {
                var cells = new List<string>();
                foreach (XElement cell in row.Elements(W + "tc"))
                {
                    var cellSb = new StringBuilder();
                    foreach (XElement p in cell.Elements(W + "p"))
                    {
                        AppendWordParagraph(cellSb, p);
                        cellSb.Append(' ');
                    }

                    cells.Add(cellSb.ToString().Trim());
                }

                sb.Append(string.Join(" ", cells));
                sb.Append('\n');
            }
        }

        private static string ExtractXlsx(string path)
        {
            using (ZipArchive zip = ZipFile.OpenRead(path))
            {
                List<string> shared = LoadSharedStrings(zip);
                var sheets = new List<Tuple<string, string>>();
                ZipArchiveEntry wb = zip.GetEntry("xl/workbook.xml");
                ZipArchiveEntry rels = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (wb == null)
                {
                    throw new InvalidOperationException("不是有效的 xlsx");
                }

                Dictionary<string, string> relMap = LoadRels(zip, rels);
                XDocument wbDoc;
                using (Stream stream = wb.Open())
                {
                    wbDoc = XDocument.Load(stream);
                }

                XElement sheetsEl = wbDoc.Root == null ? null : wbDoc.Root.Element(Main + "sheets");
                if (sheetsEl != null)
                {
                    foreach (XElement sh in sheetsEl.Elements(Main + "sheet"))
                    {
                        string name = (string)sh.Attribute("name") ?? "Sheet";
                        string rid = (string)sh.Attribute(R + "id");
                        string target;
                        if (rid != null && relMap.TryGetValue(rid, out target))
                        {
                            sheets.Add(Tuple.Create(name, NormalizeZipPath("xl/" + target.TrimStart('/'))));
                        }
                    }
                }

                var sb = new StringBuilder();
                foreach (Tuple<string, string> sheet in sheets)
                {
                    sb.AppendLine(sheet.Item1);
                    ZipArchiveEntry sheetEntry = zip.GetEntry(sheet.Item2);
                    if (sheetEntry == null)
                    {
                        continue;
                    }

                    XDocument sheetDoc;
                    using (Stream stream = sheetEntry.Open())
                    {
                        sheetDoc = XDocument.Load(stream);
                    }

                    XElement sheetData = sheetDoc.Descendants(Main + "sheetData").FirstOrDefault();
                    if (sheetData == null)
                    {
                        continue;
                    }

                    foreach (XElement row in sheetData.Elements(Main + "row"))
                    {
                        var cells = new List<string>();
                        foreach (XElement c in row.Elements(Main + "c"))
                        {
                            cells.Add(CellText(c, shared));
                        }

                        if (cells.TrueForAll(string.IsNullOrEmpty))
                        {
                            continue;
                        }

                        sb.AppendLine(string.Join("\t", cells));
                    }

                    sb.Append('\n');
                }

                return sb.ToString().TrimEnd();
            }
        }

        private static List<string> LoadSharedStrings(ZipArchive zip)
        {
            var list = new List<string>();
            ZipArchiveEntry entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return list;
            }

            XDocument doc;
            using (Stream stream = entry.Open())
            {
                doc = XDocument.Load(stream);
            }

            foreach (XElement si in doc.Descendants(Main + "si"))
            {
                list.Add(string.Concat(si.Descendants(Main + "t").Select(t => t.Value)));
            }

            return list;
        }

        private static string CellText(XElement cell, List<string> shared)
        {
            string t = (string)cell.Attribute("t");
            if (t == "s")
            {
                XElement v = cell.Element(Main + "v");
                int idx;
                if (v != null && int.TryParse(v.Value, out idx) && idx >= 0 && idx < shared.Count)
                {
                    return shared[idx];
                }

                return "";
            }

            if (t == "inlineStr")
            {
                return string.Concat(cell.Descendants(Main + "t").Select(x => x.Value));
            }

            XElement val = cell.Element(Main + "v");
            return val == null ? "" : val.Value;
        }

        private static Dictionary<string, string> LoadRels(ZipArchive zip, ZipArchiveEntry rels)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            if (rels == null)
            {
                return map;
            }

            XDocument doc;
            using (Stream stream = rels.Open())
            {
                doc = XDocument.Load(stream);
            }

            foreach (XElement rel in doc.Root.Elements(Rels + "Relationship"))
            {
                string id = (string)rel.Attribute("Id");
                string target = (string)rel.Attribute("Target");
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                {
                    map[id] = target.Replace('\\', '/');
                }
            }

            return map;
        }

        private static string NormalizeZipPath(string path)
        {
            string[] parts = path.Replace('\\', '/').Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            var stack = new List<string>();
            foreach (string part in parts)
            {
                if (part == ".")
                {
                    continue;
                }

                if (part == ".." && stack.Count > 0)
                {
                    stack.RemoveAt(stack.Count - 1);
                    continue;
                }

                if (part != "..")
                {
                    stack.Add(part);
                }
            }

            return string.Join("/", stack);
        }

        private static string ExtractPptx(string path)
        {
            using (ZipArchive zip = ZipFile.OpenRead(path))
            {
                List<ZipArchiveEntry> slides = zip.Entries
                    .Where(e => e.FullName.StartsWith("ppt/slides/slide", StringComparison.OrdinalIgnoreCase)
                        && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                        && !e.FullName.Contains("/_rels/"))
                    .OrderBy(e => e.FullName, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                var sb = new StringBuilder();
                int n = 0;
                foreach (ZipArchiveEntry slide in slides)
                {
                    n++;
                    sb.Append("第").Append(n).Append("页\n");
                    XDocument doc;
                    using (Stream stream = slide.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    foreach (XElement t in doc.Descendants(A + "t"))
                    {
                        if (!string.IsNullOrEmpty(t.Value))
                        {
                            sb.AppendLine(t.Value);
                        }
                    }

                    sb.Append('\n');
                }

                return sb.ToString().TrimEnd();
            }
        }

        private static string ExtractPdf(string path)
        {
            PdfiumNativeLoader.EnsureLoaded();
            using (PdfDocument pdf = PdfDocument.Load(path))
            {
                var sb = new StringBuilder();
                for (int i = 0; i < pdf.PageCount; i++)
                {
                    if (i > 0)
                    {
                        sb.Append('\n');
                    }

                    string page = pdf.GetPdfText(i) ?? "";
                    sb.Append(page.TrimEnd());
                    if (page.Length > 0 && !page.EndsWith("\n", StringComparison.Ordinal))
                    {
                        sb.Append('\n');
                    }
                }

                return sb.ToString().TrimEnd();
            }
        }

        private static string ExtractWordCom(string path)
        {
            object app = null;
            object doc = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Word.Application");
                if (t == null)
                {
                    throw new InvalidOperationException("无法在本机抽取 .doc，请安装 Word 或改用 docx / detailed=true");
                }

                app = Activator.CreateInstance(t);
                t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
                object docs = t.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, app, null);
                doc = docs.GetType().InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    docs,
                    new object[] { path, false, true });
                object content = doc.GetType().InvokeMember("Content", System.Reflection.BindingFlags.GetProperty, null, doc, null);
                return (content.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, content, null) as string) ?? "";
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法在本机抽取 .doc，请安装 Word 或改用 docx / detailed=true: " + ex.Message, ex);
            }
            finally
            {
                TryCom(doc, "Close", new object[] { false });
                TryCom(app, "Quit", null);
                Release(doc);
                Release(app);
            }
        }

        private static string ExtractExcelCom(string path)
        {
            object app = null;
            object book = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Excel.Application");
                if (t == null)
                {
                    throw new InvalidOperationException("无法在本机抽取该 Excel 格式，请安装 Excel 或改用 xlsx / detailed=true");
                }

                app = Activator.CreateInstance(t);
                t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
                t.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
                object books = t.InvokeMember("Workbooks", System.Reflection.BindingFlags.GetProperty, null, app, null);
                book = books.GetType().InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    books,
                    new object[] { path, 0, true });
                object sheets = book.GetType().InvokeMember("Worksheets", System.Reflection.BindingFlags.GetProperty, null, book, null);
                int count = Convert.ToInt32(sheets.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, sheets, null));
                var sb = new StringBuilder();
                for (int i = 1; i <= count; i++)
                {
                    object sheet = sheets.GetType().InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, sheets, new object[] { i });
                    string name = Convert.ToString(sheet.GetType().InvokeMember("Name", System.Reflection.BindingFlags.GetProperty, null, sheet, null));
                    sb.AppendLine(name);
                    object used = sheet.GetType().InvokeMember("UsedRange", System.Reflection.BindingFlags.GetProperty, null, sheet, null);
                    object val = used.GetType().InvokeMember("Value2", System.Reflection.BindingFlags.GetProperty, null, used, null);
                    AppendExcelValue(sb, val);
                    sb.Append('\n');
                    Release(used);
                    Release(sheet);
                }

                return sb.ToString().TrimEnd();
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法在本机抽取该 Excel 格式，请安装 Excel 或改用 xlsx / detailed=true: " + ex.Message, ex);
            }
            finally
            {
                TryCom(book, "Close", new object[] { false });
                TryCom(app, "Quit", null);
                Release(book);
                Release(app);
            }
        }

        private static void AppendExcelValue(StringBuilder sb, object val)
        {
            var arr = val as object[,];
            if (arr != null)
            {
                int r0 = arr.GetLowerBound(0);
                int r1 = arr.GetUpperBound(0);
                int c0 = arr.GetLowerBound(1);
                int c1 = arr.GetUpperBound(1);
                for (int r = r0; r <= r1; r++)
                {
                    var cells = new List<string>();
                    for (int c = c0; c <= c1; c++)
                    {
                        object cell = arr[r, c];
                        cells.Add(cell == null ? "" : Convert.ToString(cell));
                    }

                    if (!cells.TrueForAll(string.IsNullOrEmpty))
                    {
                        sb.AppendLine(string.Join("\t", cells));
                    }
                }

                return;
            }

            if (val != null)
            {
                sb.AppendLine(Convert.ToString(val));
            }
        }

        private static string ExtractPptCom(string path)
        {
            object app = null;
            object pres = null;
            try
            {
                Type t = Type.GetTypeFromProgID("PowerPoint.Application");
                if (t == null)
                {
                    throw new InvalidOperationException("无法在本机抽取 .ppt，请安装 PowerPoint 或改用 pptx / detailed=true");
                }

                app = Activator.CreateInstance(t);
                object presentations = t.InvokeMember("Presentations", System.Reflection.BindingFlags.GetProperty, null, app, null);
                pres = presentations.GetType().InvokeMember(
                    "Open",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    presentations,
                    new object[] { path, true, false, false });
                object slides = pres.GetType().InvokeMember("Slides", System.Reflection.BindingFlags.GetProperty, null, pres, null);
                int count = Convert.ToInt32(slides.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, slides, null));
                var sb = new StringBuilder();
                for (int i = 1; i <= count; i++)
                {
                    object slide = slides.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, slides, new object[] { i });
                    sb.Append("第").Append(i).Append("页\n");
                    object shapes = slide.GetType().InvokeMember("Shapes", System.Reflection.BindingFlags.GetProperty, null, slide, null);
                    int sc = Convert.ToInt32(shapes.GetType().InvokeMember("Count", System.Reflection.BindingFlags.GetProperty, null, shapes, null));
                    for (int s = 1; s <= sc; s++)
                    {
                        object shape = shapes.GetType().InvokeMember("Item", System.Reflection.BindingFlags.InvokeMethod, null, shapes, new object[] { s });
                        try
                        {
                            object tf = shape.GetType().InvokeMember("HasTextFrame", System.Reflection.BindingFlags.GetProperty, null, shape, null);
                            if (Convert.ToInt32(tf) == -1)
                            {
                                object frame = shape.GetType().InvokeMember("TextFrame", System.Reflection.BindingFlags.GetProperty, null, shape, null);
                                object range = frame.GetType().InvokeMember("TextRange", System.Reflection.BindingFlags.GetProperty, null, frame, null);
                                string text = range.GetType().InvokeMember("Text", System.Reflection.BindingFlags.GetProperty, null, range, null) as string;
                                if (!string.IsNullOrWhiteSpace(text))
                                {
                                    sb.AppendLine(text.Trim());
                                }

                                Release(range);
                                Release(frame);
                            }
                        }
                        catch
                        {
                            /* 无文本框 */
                        }

                        Release(shape);
                    }

                    sb.Append('\n');
                    Release(shapes);
                    Release(slide);
                }

                return sb.ToString().TrimEnd();
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("无法在本机抽取 .ppt，请安装 PowerPoint 或改用 pptx / detailed=true: " + ex.Message, ex);
            }
            finally
            {
                TryCom(pres, "Close", null);
                TryCom(app, "Quit", null);
                Release(pres);
                Release(app);
            }
        }

        private static void TryCom(object obj, string method, object[] args)
        {
            if (obj == null)
            {
                return;
            }

            try
            {
                obj.GetType().InvokeMember(method, System.Reflection.BindingFlags.InvokeMethod, null, obj, args);
            }
            catch
            {
                /* ignore */
            }
        }

        private static void Release(object obj)
        {
            if (obj == null)
            {
                return;
            }

            try
            {
                Marshal.ReleaseComObject(obj);
            }
            catch
            {
                /* ignore */
            }
        }
    }
}
