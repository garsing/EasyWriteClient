using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>Excel / PPT 检查点短套：经 Desktop 管道调 F_*，COM 只用来关稿和造「不建副本」的改动。</summary>
    internal sealed class OfficeCkptSession
    {
        public OfficeCkptSession(string host)
        {
            Host = NormalizeHost(host);
            App = Host == "ppt" ? "powerpoint" : Host;
            ExpectedKind = Host;
        }

        public string Host { get; }

        public string App { get; }

        public string ExpectedKind { get; }

        public bool IsExcelFamily
        {
            get { return Host == "excel" || Host == "et"; }
        }

        public bool IsEt
        {
            get { return Host == "et"; }
        }

        public bool IsWpp
        {
            get { return Host == "wpp"; }
        }

        public string ProcName
        {
            get
            {
                if (Host == "excel")
                {
                    return "EXCEL";
                }

                if (Host == "et")
                {
                    return "et";
                }

                if (Host == "wpp")
                {
                    return "wpp";
                }

                return "POWERPNT";
            }
        }

        public void SweepLeftovers()
        {
            try
            {
                if (IsExcelFamily)
                {
                    SweepBooks(IsEt, "ew-xlk-");
                }
                else
                {
                    SweepPresentations(IsWpp, "ew-pk-");
                }
            }
            catch
            {
            }

            try
            {
                string prefix = IsExcelFamily ? "ew-xlk-" : "ew-pk-";
                string ext = IsExcelFamily ? ".xlsx" : ".pptx";
                foreach (string file in Directory.GetFiles(Path.GetTempPath(), prefix + "*" + ext))
                {
                    try
                    {
                        File.Delete(file);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
        }

        public string NewPath()
        {
            string ext = IsExcelFamily ? ".xlsx" : ".pptx";
            string tag = IsExcelFamily ? "ew-xlk-" : "ew-pk-";
            return Path.Combine(Path.GetTempPath(), tag + Guid.NewGuid().ToString("N") + ext);
        }

        public ToolResult Open(string path)
        {
            ToolResult created = DesktopToolPipe.Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = path,
                ["create_blank"] = true,
                ["app"] = App
            });
            if (created.Success)
            {
                return created;
            }

            if (!TrySeedBlankFile(path))
            {
                return created;
            }

            return DesktopToolPipe.Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = path,
                ["create_blank"] = false,
                ["app"] = App
            });
        }

        public ToolResult ListOperations(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call("F_list_document_operations", args);
        }

        public ToolResult RestoreCheckpoint(string channelId, string timestamp)
        {
            var args = new Dictionary<string, object>();
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (timestamp != null)
            {
                args["timestamp"] = timestamp;
            }

            return DesktopToolPipe.Call("F_restore_document_checkpoint", args);
        }

        public ToolResult Overview(string channelId)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            return DesktopToolPipe.Call(
                IsExcelFamily ? "F_get_workbook_content" : "F_get_presentation_content",
                args);
        }

        public ToolResult AddSheet(string channelId, string name)
        {
            return DesktopToolPipe.Call("F_manage_excel_sheet", new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["action"] = "add",
                ["name"] = name
            });
        }

        public ToolResult AddSlide(string channelId)
        {
            return DesktopToolPipe.Call("F_manage_ppt_slide", new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["action"] = "add"
            });
        }

        public ToolResult MutatingContractFail(string channelId)
        {
            string tool = IsExcelFamily ? "F_manage_excel_sheet" : "F_manage_ppt_slide";
            return DesktopToolPipe.Call(tool, new Dictionary<string, object>
            {
                ["channel_id"] = channelId
            });
        }

        public bool OverviewHas(string channelId, string needle)
        {
            ToolResult r = Overview(channelId);
            if (r == null || !r.Success)
            {
                return false;
            }

            return CkptCaseUtil.BlobHas(r, needle);
        }

        public int SlideCount(string channelId)
        {
            for (int i = 0; i < 4; i++)
            {
                int n = WdSession.IntOf(Overview(channelId), "slide_count");
                if (n > 0)
                {
                    return n;
                }

                System.Threading.Thread.Sleep(200);
            }

            return 0;
        }

        public bool ComAddSheet(string path, string name)
        {
            for (int i = 0; i < 8; i++)
            {
                if (IsEt ? EtAddSheet(path, name) : ExcelAddSheet(path, name))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        public bool ComSelectCell(string path, string address)
        {
            for (int i = 0; i < 8; i++)
            {
                if (SelectCellLate(path, address, IsEt))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        public bool TryReadActiveCell(string path, out string address)
        {
            address = "";
            string found = "";
            bool ok = ReadActiveCellLate(path, IsEt, out found);
            address = found ?? "";
            return ok && !string.IsNullOrEmpty(address);
        }

        public bool ComGotoSlide(string path, int index)
        {
            for (int i = 0; i < 8; i++)
            {
                if (GotoSlideLate(path, index, IsWpp))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        public bool TryReadActiveSlide(string path, out int index)
        {
            index = 0;
            int found = 0;
            bool ok = ReadActiveSlideLate(path, IsWpp, out found);
            index = found;
            return ok && index >= 1;
        }

        public static string NormalizeCell(string address)
        {
            if (string.IsNullOrEmpty(address))
            {
                return "";
            }

            return address.Replace("$", "").Trim();
        }

        public bool ComAddSlide(string path)
        {
            for (int i = 0; i < 8; i++)
            {
                if (IsWpp ? WppAddSlide(path) : PptAddSlide(path))
                {
                    return true;
                }

                System.Threading.Thread.Sleep(200);
            }

            return false;
        }

        private bool TrySeedBlankFile(string path)
        {
            try
            {
                if (IsExcelFamily)
                {
                    return SeedWorkbook(path, IsEt) || SeedMinimalXlsx(path);
                }

                return SeedPresentation(path, IsWpp);
            }
            catch
            {
                return false;
            }
        }

        private static bool SeedWorkbook(string path, bool et)
        {
            try
            {
                object app;
                string err;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out err, true)
                    : ExcelApplicationResolver.TryResolve(out app, out err, true);
                if (!ok)
                {
                    return false;
                }

                Late.Set(app, "DisplayAlerts", false);
                object books = Late.Get(app, "Workbooks");
                object book = Late.Call(books, "Add");
                try
                {
                    Late.Call(book, "SaveAs", path, 51);
                }
                finally
                {
                    try
                    {
                        Late.Call(book, "Close", false);
                    }
                    catch
                    {
                    }
                }

                return File.Exists(path) && new FileInfo(path).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool SeedPresentation(string path, bool wpp)
        {
            object app;
            string err;
            bool ok = wpp
                ? WppApplicationResolver.TryResolve(out app, out err, true)
                : PowerPointApplicationResolver.TryResolve(out app, out err, true);
            if (!ok)
            {
                return false;
            }

            object col = Late.Get(app, "Presentations");
            object pres = Late.Call(col, "Add");
            try
            {
                try
                {
                    Late.Call(pres, "SaveAs", path, 24);
                }
                catch
                {
                    Late.Call(pres, "SaveAs", path);
                }
            }
            finally
            {
                try
                {
                    Late.Call(pres, "Close");
                }
                catch
                {
                }
            }

            return File.Exists(path) && new FileInfo(path).Length > 0;
        }

        private static bool SeedMinimalXlsx(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }

                using (var fs = File.Create(path))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
                {
                    WriteZip(zip, "[Content_Types].xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                        + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                        + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                        + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                        + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                        + "</Types>");
                    WriteZip(zip, "_rels/.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                        + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                        + "</Relationships>");
                    WriteZip(zip, "xl/workbook.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                        + "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                        + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                    WriteZip(zip, "xl/_rels/workbook.xml.rels",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                        + "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                        + "</Relationships>");
                    WriteZip(zip, "xl/worksheets/sheet1.xml",
                        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + "<sheetData/></worksheet>");
                }

                return File.Exists(path) && new FileInfo(path).Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static void WriteZip(System.IO.Compression.ZipArchive zip, string entry, string xml)
        {
            using (var w = new StreamWriter(zip.CreateEntry(entry).Open()))
            {
                w.Write(xml);
            }
        }

        public void CloseByPath(string path, string channelId = null)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            try
            {
                if (!string.IsNullOrEmpty(channelId))
                {
                    DesktopToolPipe.Call("F_close_document", new Dictionary<string, object>
                    {
                        ["channel_id"] = channelId
                    });
                }
                else if (IsExcelFamily)
                {
                    CloseLate(path, IsEt);
                }
                else
                {
                    ClosePresLate(path, IsWpp);
                }
            }
            catch
            {
            }

            try
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        public static string NormalizeExcelHost(string host)
        {
            if (string.Equals(host, "et", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase))
            {
                return "et";
            }

            return "excel";
        }

        public static string NormalizePptHost(string host)
        {
            if (string.Equals(host, "wpp", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wps", StringComparison.OrdinalIgnoreCase))
            {
                return "wpp";
            }

            return "ppt";
        }

        private static string NormalizeHost(string host)
        {
            if (string.Equals(host, "et", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "wpp", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "ppt", StringComparison.OrdinalIgnoreCase)
                || string.Equals(host, "excel", StringComparison.OrdinalIgnoreCase))
            {
                return host.ToLowerInvariant();
            }

            return host ?? "";
        }

        private static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\', '/'),
                    Path.GetFullPath(b).TrimEnd('\\', '/'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        private static bool ExcelAddSheet(string path, string name)
        {
            return AddSheetLate(path, name, false);
        }

        private static bool EtAddSheet(string path, string name)
        {
            return AddSheetLate(path, name, true);
        }

        private static bool PptAddSlide(string path)
        {
            return AddSlideLate(path, false);
        }

        private static bool WppAddSlide(string path)
        {
            return AddSlideLate(path, true);
        }

        private static bool AddSheetLate(string path, string name, bool et)
        {
            try
            {
                object app;
                string err;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out err, false)
                    : ExcelApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return false;
                }

                object book = FindLateBook(app, "Workbooks", path);
                if (book == null)
                {
                    return false;
                }

                object sheets = Late.Get(book, "Worksheets");
                object ws = Late.Call(sheets, "Add");
                Late.Set(ws, "Name", name);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void CloseLate(string path, bool et)
        {
            try
            {
                object app;
                string err;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out err, false)
                    : ExcelApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return;
                }

                object book = FindLateBook(app, "Workbooks", path);
                if (book != null)
                {
                    Late.Call(book, "Close", false);
                }
            }
            catch
            {
            }
        }

        private static bool SelectCellLate(string path, string address, bool et)
        {
            try
            {
                object book = FindWorkbook(path, et);
                if (book == null)
                {
                    return false;
                }

                object sheet = Late.Get(book, "ActiveSheet");
                object range = Late.Get(sheet, "Range", address);
                Late.Call(range, "Select");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool ReadActiveCellLate(string path, bool et, out string address)
        {
            address = "";
            try
            {
                object book = FindWorkbook(path, et);
                if (book == null)
                {
                    return false;
                }

                object app = Late.Get(book, "Application");
                object cell = Late.Get(app, "ActiveCell");
                address = Convert.ToString(Late.Get(cell, "Address") ?? "");
                return !string.IsNullOrEmpty(address);
            }
            catch
            {
                return false;
            }
        }

        private static void SweepBooks(bool et, string namePrefix)
        {
            object app;
            string err;
            bool ok = et
                ? EtApplicationResolver.TryResolve(out app, out err, false)
                : ExcelApplicationResolver.TryResolve(out app, out err, false);
            if (!ok)
            {
                return;
            }

            try
            {
                Late.Set(app, "DisplayAlerts", false);
            }
            catch
            {
            }

            object col = Late.Get(app, "Workbooks");
            int n = Convert.ToInt32(Late.Get(col, "Count"));
            for (int i = n; i >= 1; i--)
            {
                object item = Late.Item(col, i);
                string full = Convert.ToString(Late.Get(item, "FullName") ?? "");
                string name = Convert.ToString(Late.Get(item, "Name") ?? "");
                if (ContainsIgnoreCase(full, namePrefix) || ContainsIgnoreCase(name, namePrefix))
                {
                    try
                    {
                        Late.Call(item, "Close", false);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static void SweepPresentations(bool wpp, string namePrefix)
        {
            object app;
            string err;
            bool ok = wpp
                ? WppApplicationResolver.TryResolve(out app, out err, false)
                : PowerPointApplicationResolver.TryResolve(out app, out err, false);
            if (!ok)
            {
                return;
            }

            object col = Late.Get(app, "Presentations");
            int n = Convert.ToInt32(Late.Get(col, "Count"));
            for (int i = n; i >= 1; i--)
            {
                object item = Late.Item(col, i);
                string full = Convert.ToString(Late.Get(item, "FullName") ?? "");
                string name = Convert.ToString(Late.Get(item, "Name") ?? "");
                if (ContainsIgnoreCase(full, namePrefix) || ContainsIgnoreCase(name, namePrefix))
                {
                    try
                    {
                        Late.Call(item, "Close");
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static bool ContainsIgnoreCase(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static object FindWorkbook(string path, bool et)
        {
            object app;
            string err;
            bool ok = et
                ? EtApplicationResolver.TryResolve(out app, out err, false)
                : ExcelApplicationResolver.TryResolve(out app, out err, false);
            if (!ok)
            {
                return null;
            }

            return FindLateBook(app, "Workbooks", path);
        }

        private static bool GotoSlideLate(string path, int index, bool wpp)
        {
            try
            {
                object pres = FindPresentation(path, wpp);
                if (pres == null)
                {
                    return false;
                }

                try
                {
                    object view = Late.Get(Late.Get(Late.Get(pres, "Application"), "ActiveWindow"), "View");
                    Late.Call(view, "GotoSlide", index);
                    return true;
                }
                catch
                {
                }

                object slides = Late.Get(pres, "Slides");
                object slide = Late.Item(slides, index);
                Late.Call(slide, "Select");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool ReadActiveSlideLate(string path, bool wpp, out int index)
        {
            index = 0;
            try
            {
                object pres = FindPresentation(path, wpp);
                if (pres == null)
                {
                    return false;
                }

                object app = Late.Get(pres, "Application");
                object view = Late.Get(Late.Get(app, "ActiveWindow"), "View");
                object slide = Late.Get(view, "Slide");
                index = Convert.ToInt32(Late.Get(slide, "SlideIndex"));
                return index >= 1;
            }
            catch
            {
                return false;
            }
        }

        private static object FindPresentation(string path, bool wpp)
        {
            object app;
            string err;
            bool ok = wpp
                ? WppApplicationResolver.TryResolve(out app, out err, false)
                : PowerPointApplicationResolver.TryResolve(out app, out err, false);
            if (!ok)
            {
                return null;
            }

            return FindLateBook(app, "Presentations", path);
        }

        private static bool AddSlideLate(string path, bool wpp)
        {
            try
            {
                object app;
                string err;
                bool ok = wpp
                    ? WppApplicationResolver.TryResolve(out app, out err, false)
                    : PowerPointApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return false;
                }

                object pres = FindLateBook(app, "Presentations", path);
                if (pres == null)
                {
                    return false;
                }

                object slides = Late.Get(pres, "Slides");
                int n = Convert.ToInt32(Late.Get(slides, "Count"));
                try
                {
                    Late.Call(slides, "Add", n + 1, 2);
                    return true;
                }
                catch
                {
                }

                object layouts = Late.Get(Late.Get(pres, "SlideMaster"), "CustomLayouts");
                object layout = Late.Item(layouts, 1);
                Late.Call(slides, "AddSlide", n + 1, layout);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static void ClosePresLate(string path, bool wpp)
        {
            try
            {
                object app;
                string err;
                bool ok = wpp
                    ? WppApplicationResolver.TryResolve(out app, out err, false)
                    : PowerPointApplicationResolver.TryResolve(out app, out err, false);
                if (!ok)
                {
                    return;
                }

                object pres = FindLateBook(app, "Presentations", path);
                if (pres != null)
                {
                    Late.Call(pres, "Close");
                }
            }
            catch
            {
            }
        }

        private static object FindLateBook(object app, string collectionName, string path)
        {
            object col = Late.Get(app, collectionName);
            int n = Convert.ToInt32(Late.Get(col, "Count"));
            for (int i = 1; i <= n; i++)
            {
                object item = Late.Item(col, i);
                string full = Convert.ToString(Late.Get(item, "FullName") ?? "");
                string name = Convert.ToString(Late.Get(item, "Name") ?? "");
                if (SamePath(full, path)
                    || string.Equals(name, Path.GetFileName(path), StringComparison.OrdinalIgnoreCase))
                {
                    return item;
                }
            }

            return null;
        }

        private static class Late
        {
            public static object Get(object target, string name, params object[] args)
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.GetProperty,
                    null,
                    target,
                    args == null || args.Length == 0 ? null : args);
            }

            public static void Set(object target, string name, object value)
            {
                target.GetType().InvokeMember(
                    name,
                    BindingFlags.SetProperty,
                    null,
                    target,
                    new[] { value });
            }

            public static object Call(object target, string name, params object[] args)
            {
                return target.GetType().InvokeMember(
                    name,
                    BindingFlags.InvokeMethod,
                    null,
                    target,
                    args == null || args.Length == 0 ? null : args);
            }

            public static object Item(object target, object key)
            {
                try
                {
                    return target.GetType().InvokeMember(
                        "Item",
                        BindingFlags.GetProperty,
                        null,
                        target,
                        new[] { key });
                }
                catch
                {
                    return target.GetType().InvokeMember(
                        "Item",
                        BindingFlags.InvokeMethod,
                        null,
                        target,
                        new[] { key });
                }
            }
        }
    }
}
