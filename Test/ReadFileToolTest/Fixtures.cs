using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace ReadFileToolTest
{
    internal sealed class Fixtures : IDisposable
    {
        public string Dir { get; }
        public string NotesMd { get; }
        public string Png { get; }
        public string Zip { get; }
        public string ShortDocx { get; }
        public string LongDocx { get; }
        public string Xlsx { get; }
        public string Pptx { get; }
        public string TextPdf { get; }
        public string ScanPdf { get; }
        public string BadPdf { get; }
        public string EmptyFile { get; }
        public string RealPdf { get; }
        public string LegacyDoc { get; }
        public string LegacyDocError { get; }

        public Fixtures()
        {
            Dir = Path.Combine(Path.GetTempPath(), "ew-rf-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Dir);
            NotesMd = Path.Combine(Dir, "notes.md");
            File.WriteAllText(NotesMd, "# 标题\nhello read_file\n", new UTF8Encoding(false));
            Png = Path.Combine(Dir, "dot.png");
            File.WriteAllBytes(Png, Convert.FromBase64String(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            Zip = Path.Combine(Dir, "x.zip");
            using (var z = ZipFile.Open(Zip, ZipArchiveMode.Create))
            {
                ZipArchiveEntry e = z.CreateEntry("a.txt");
                using (var w = new StreamWriter(e.Open()))
                {
                    w.Write("z");
                }
            }

            ShortDocx = Path.Combine(Dir, "short.docx");
            WriteDocx(ShortDocx, "短文标题微专业", "格甲", "格乙");
            LongDocx = Path.Combine(Dir, "long.docx");
            WriteDocx(LongDocx, new string('甲', 8000) + new string('乙', 5000), null, null);
            Xlsx = Path.Combine(Dir, "two.xlsx");
            WriteXlsx(Xlsx);
            Pptx = Path.Combine(Dir, "two.pptx");
            WritePptx(Pptx);
            TextPdf = Path.Combine(Dir, "text.pdf");
            WritePdf(TextPdf, "BT /F1 12 Tf 72 720 Td (HelloMicro) Tj ET");
            ScanPdf = Path.Combine(Dir, "scan.pdf");
            WritePdf(ScanPdf, "");
            BadPdf = Path.Combine(Dir, "bad.pdf");
            File.WriteAllText(BadPdf, "%PDF-1.4\nnot-a-pdf");
            EmptyFile = Path.Combine(Dir, "zero.pdf");
            File.WriteAllBytes(EmptyFile, Array.Empty<byte>());

            string[] candidates =
            {
                @"D:\documents\易写\易写产品试用\2. 体验素材\附件2微专业-智能数据工程与标注技术v6.0.pdf",
                Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory,
                    "..", "..", "..", "..", "..",
                    "WordAddInBackend", "main-server", "project", "tests", "fixtures", "sample_2page.pdf")
            };
            foreach (string c in candidates)
            {
                string full = Path.GetFullPath(c);
                if (File.Exists(full))
                {
                    RealPdf = full;
                    break;
                }
            }

            LegacyDoc = Path.Combine(Dir, "legacy.doc");
            LegacyDocError = TryWriteLegacyDoc(LegacyDoc);
            if (!string.IsNullOrEmpty(LegacyDocError))
            {
                LegacyDoc = null;
            }
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(Dir, true);
            }
            catch
            {
                /* keep temp on lock */
            }
        }

        private static string TryWriteLegacyDoc(string path)
        {
            object app = null;
            object doc = null;
            try
            {
                Type t = Type.GetTypeFromProgID("Word.Application");
                if (t == null)
                {
                    return "本机没有 Word.Application";
                }

                app = Activator.CreateInstance(t);
                t.InvokeMember("Visible", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { false });
                t.InvokeMember("DisplayAlerts", System.Reflection.BindingFlags.SetProperty, null, app, new object[] { 0 });
                object docs = t.InvokeMember("Documents", System.Reflection.BindingFlags.GetProperty, null, app, null);
                doc = docs.GetType().InvokeMember(
                    "Add",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    docs,
                    null);
                object content = doc.GetType().InvokeMember("Content", System.Reflection.BindingFlags.GetProperty, null, doc, null);
                content.GetType().InvokeMember("Text", System.Reflection.BindingFlags.SetProperty, null, content, new object[] { "老doc正文微专业" });
                const int wdFormatDocument = 0;
                doc.GetType().InvokeMember(
                    "SaveAs",
                    System.Reflection.BindingFlags.InvokeMethod,
                    null,
                    doc,
                    new object[] { path, wdFormatDocument });
                return File.Exists(path) ? null : "SaveAs 后没有 .doc";
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        doc.GetType().InvokeMember("Close", System.Reflection.BindingFlags.InvokeMethod, null, doc, new object[] { false });
                    }
                }
                catch
                {
                    /* ignore */
                }

                try
                {
                    if (app != null)
                    {
                        app.GetType().InvokeMember("Quit", System.Reflection.BindingFlags.InvokeMethod, null, app, new object[] { false });
                    }
                }
                catch
                {
                    /* ignore */
                }

                if (doc != null)
                {
                    Marshal.ReleaseComObject(doc);
                }

                if (app != null)
                {
                    Marshal.ReleaseComObject(app);
                }
            }
        }

        private static void WriteDocx(string path, string para, string cellA, string cellB)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>"
                    + "</Types>");
                WriteEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>"
                    + "</Relationships>");
                var body = new StringBuilder();
                body.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
                body.Append("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
                body.Append("<w:p><w:r><w:t>").Append(Xml(para)).Append("</w:t></w:r></w:p>");
                if (cellA != null)
                {
                    body.Append("<w:tbl><w:tr>");
                    body.Append("<w:tc><w:p><w:r><w:t>").Append(Xml(cellA)).Append("</w:t></w:r></w:p></w:tc>");
                    body.Append("<w:tc><w:p><w:r><w:t>").Append(Xml(cellB)).Append("</w:t></w:r></w:p></w:tc>");
                    body.Append("</w:tr></w:tbl>");
                }

                body.Append("</w:body></w:document>");
                WriteEntry(zip, "word/document.xml", body.ToString());
            }
        }

        private static void WriteXlsx(string path)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                    + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                    + "<Override PartName=\"/xl/worksheets/sheet2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                    + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                    + "</Types>");
                WriteEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                    + "</Relationships>");
                WriteEntry(zip, "xl/_rels/workbook.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                    + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
                    + "<Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>"
                    + "</Relationships>");
                WriteEntry(zip, "xl/workbook.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "<sheets><sheet name=\"收入\" sheetId=\"1\" r:id=\"rId1\"/><sheet name=\"成本\" sheetId=\"2\" r:id=\"rId2\"/></sheets></workbook>");
                WriteEntry(zip, "xl/sharedStrings.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">"
                    + "<si><t>单元格甲</t></si><si><t>单元格乙</t></si></sst>");
                WriteEntry(zip, "xl/worksheets/sheet1.xml", SheetXml(0));
                WriteEntry(zip, "xl/worksheets/sheet2.xml", SheetXml(1));
            }
        }

        private static string SheetXml(int sharedIndex)
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                + "<sheetData><row r=\"1\"><c r=\"A1\" t=\"s\"><v>" + sharedIndex + "</v></c></row></sheetData>"
                + "</worksheet>";
        }

        private static void WritePptx(string path)
        {
            using (var zip = ZipFile.Open(path, ZipArchiveMode.Create))
            {
                WriteEntry(zip, "[Content_Types].xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                    + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                    + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                    + "<Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>"
                    + "<Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
                    + "<Override PartName=\"/ppt/slides/slide2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>"
                    + "</Types>");
                WriteEntry(zip, "_rels/.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"ppt/presentation.xml\"/>"
                    + "</Relationships>");
                WriteEntry(zip, "ppt/_rels/presentation.xml.rels",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                    + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                    + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>"
                    + "<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide2.xml\"/>"
                    + "</Relationships>");
                WriteEntry(zip, "ppt/presentation.xml",
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                    + "<p:sldIdLst><p:sldId id=\"256\" r:id=\"rId1\"/><p:sldId id=\"257\" r:id=\"rId2\"/></p:sldIdLst></p:presentation>");
                WriteEntry(zip, "ppt/slides/slide1.xml", SlideXml("第一页标题"));
                WriteEntry(zip, "ppt/slides/slide2.xml", SlideXml("第二页正文"));
            }
        }

        private static string SlideXml(string text)
        {
            return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                + "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
                + "<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>"
                + "<p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id=\"2\" name=\"t\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
                + "<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>" + Xml(text) + "</a:t></a:r></a:p></p:txBody>"
                + "</p:sp></p:spTree></p:cSld></p:sld>";
        }

        private static void WritePdf(string path, string contentStream)
        {
            contentStream = contentStream ?? "";
            string stream = contentStream + "\n";
            var sb = new StringBuilder();
            sb.Append("%PDF-1.1\n");
            int[] offs = new int[6];
            offs[1] = sb.Length;
            sb.Append("1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj\n");
            offs[2] = sb.Length;
            sb.Append("2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj\n");
            offs[3] = sb.Length;
            sb.Append("3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj\n");
            offs[4] = sb.Length;
            sb.Append("4 0 obj << /Length ").Append(Encoding.ASCII.GetByteCount(stream)).Append(" >> stream\n");
            sb.Append(stream);
            sb.Append("endstream endobj\n");
            offs[5] = sb.Length;
            sb.Append("5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj\n");
            int xref = sb.Length;
            sb.Append("xref\n0 6\n0000000000 65535 f \n");
            for (int i = 1; i <= 5; i++)
            {
                sb.Append(offs[i].ToString("D10")).Append(" 00000 n \n");
            }

            sb.Append("trailer << /Size 6 /Root 1 0 R >>\nstartxref\n").Append(xref).Append("\n%%EOF\n");
            File.WriteAllText(path, sb.ToString(), Encoding.ASCII);
        }

        private static void WriteEntry(ZipArchive zip, string name, string xml)
        {
            ZipArchiveEntry e = zip.CreateEntry(name, CompressionLevel.Fastest);
            using (var w = new StreamWriter(e.Open(), new UTF8Encoding(false)))
            {
                w.Write(xml);
            }
        }

        private static string Xml(string s)
        {
            return (s ?? "").Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;");
        }
    }
}
