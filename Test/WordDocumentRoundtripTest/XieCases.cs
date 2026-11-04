using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// F_excel_image action=export 落到 curr_docs/{簿名}_excel 或 _et。前缀 xie-。正式 20。环境条不在这里。
    /// </summary>
    internal static class XieCases
    {
        public const int Target = 20;

        public static string IdFilter;

        private static readonly string[] MainIds =
        {
            "xie-list-nopath", "xie-sheet", "xie-id", "xie-msg", "xie-abs", "xie-png",
            "xie-dir", "xie-no-session", "xie-noclear", "xie-book", "xie-ckpt",
            "xie-keep-gone", "xie-reuse", "xie-ws"
        };

        private static readonly string[] EmptyIds =
        {
            "xie-empty", "xie-bad-sheet", "xie-bad-id"
        };

        private static readonly Regex ImageIdRx = new Regex(@"^I_[A-Za-z0-9]{8}$", RegexOptions.Compiled);

        private const string TinyPng =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

        public static void Run(TestRun run, string host)
        {
            var wanted = new List<string>();
            AddWanted(wanted, MainIds);
            AddWanted(wanted, EmptyIds);
            if (Want("xie-bad-fmt"))
            {
                wanted.Add("xie-bad-fmt");
            }

            if (Want("xie-word-ch"))
            {
                wanted.Add("xie-word-ch");
            }

            if (Want("xie-chart"))
            {
                wanted.Add("xie-chart");
            }

            run.TotalTarget = wanted.Count;
            string root;
            string sessionDir;
            string workspaceError;
            if (!TryWorkspace(out root, out sessionDir, out workspaceError))
            {
                FailIds(run, wanted, workspaceError);
                return;
            }

            string curr = Path.Combine(root.Replace('/', '\\'), "curr_docs");
            string sessions = Path.Combine(root.Replace('/', '\\'), "sessions");
            string stamp = DateTime.UtcNow.ToString("yyyyMMddHHmmssfff");
            bool et = host == "et";
            string suffix = et ? CurrDocsImagePaths.EtSuffix : CurrDocsImagePaths.ExcelSuffix;
            Console.WriteLine("工作区 " + root);
            var session = new OfficeCkptSession(host);
            Sweep(session);

            if (WantsAny(MainIds))
            {
                RunMain(run, session, et, curr, sessions, sessionDir, stamp, suffix);
            }

            if (Want("xie-bad-fmt") || WantsAny(EmptyIds))
            {
                RunEmpty(run, session, curr, stamp, suffix);
            }

            if (Want("xie-word-ch"))
            {
                RunWord(run, stamp);
            }

            if (Want("xie-chart"))
            {
                RunChart(run, session, et, curr, stamp, suffix);
            }
        }

        private static void RunMain(
            TestRun run,
            OfficeCkptSession session,
            bool et,
            string curr,
            string sessions,
            string sessionDir,
            string stamp,
            string suffix)
        {
            string png = WritePng(stamp);
            string path = Path.Combine(Path.GetTempPath(), "ew-xie-" + stamp + ".xlsx");
            string channel = null;
            try
            {
                Console.WriteLine("打开工作簿 " + path);
                ToolResult opened = session.Open(path);
                if (!opened.Success)
                {
                    FailIds(run, Wanted(MainIds), "开簿失败: " + (opened.Error ?? ""));
                    return;
                }

                channel = WdSession.Str(opened, "channel_id");
                ToolResult overview = session.Overview(channel);
                string bookName = WdSession.Str(overview, "name");
                string firstSheet = FirstSheet(overview);
                if (!overview.Success || string.IsNullOrEmpty(bookName) || string.IsNullOrEmpty(firstSheet))
                {
                    FailIds(run, Wanted(MainIds), "读不到簿名或第一张表: " + (overview.Error ?? ""));
                    return;
                }

                string stem = CurrDocsImagePaths.SanitizeStem(bookName);
                string imageDir = Path.Combine(curr, stem + "_" + suffix);
                Console.WriteLine("簿 " + bookName + "，加表并贴两张图");
                ToolResult added = session.AddSheet(channel, "分析");
                ToolResult redIns = added.Success
                    ? Image(channel, "insert", firstSheet, "B2", png, null, null)
                    : added;
                string redId = WdSession.Str(redIns, "image_id");
                ToolResult blueIns = redIns.Success
                    ? Image(channel, "insert", "分析", "C3", png, null, null)
                    : redIns;
                string blueId = WdSession.Str(blueIns, "image_id");
                if (!added.Success || !redIns.Success || !blueIns.Success
                    || !ImageIdRx.IsMatch(redId) || !ImageIdRx.IsMatch(blueId))
                {
                    FailIds(run, Wanted(MainIds), added.Error ?? redIns.Error ?? blueIns.Error ?? "铺图失败");
                    return;
                }

                ToolResult listed = Image(channel, "list", null, null, null, null, null);
                List<Pic> listedPics = Pictures(listed);
                Ok(run, "xie-list-nopath",
                    listed.Success
                        && WdSession.IntOf(listed, "images_count") == 2
                        && listedPics.Count == 2
                        && !JsonHasPath(listed)
                        && WdSession.Str(listed, "message") != "抽取成功"
                        && !Directory.Exists(imageDir),
                    listed.Success ? "count=" + WdSession.IntOf(listed, "images_count") + " 目录=" + imageDir : (listed.Error ?? ""));

                Directory.CreateDirectory(imageDir);
                string note = Path.Combine(imageDir, "note.txt");
                File.WriteAllText(note, "xie-note");
                string decoy = Path.Combine(imageDir, "I_keepme.png");
                File.WriteAllBytes(decoy, Convert.FromBase64String(TinyPng));
                string wordDir = Path.Combine(curr, stem + "_word");
                string pptDir = Path.Combine(curr, stem + "_ppt");
                Directory.CreateDirectory(wordDir);
                Directory.CreateDirectory(pptDir);
                string wordNote = Path.Combine(wordDir, "note.txt");
                string pptNote = Path.Combine(pptDir, "note.txt");
                File.WriteAllText(wordNote, "xie-word");
                File.WriteAllText(pptNote, "xie-ppt");

                DirtyCell(et, path);
                string fullBefore;
                bool savedBefore;
                bool sawBefore = TryReadBook(et, path, out _, out fullBefore, out savedBefore);
                string opsBefore = CkptCaseUtil.DataText(session.ListOperations(channel));

                Console.WriteLine("只导出「分析」");
                ToolResult sheetExport = Image(channel, "export", "分析", null, null, null, null);
                List<Pic> sheetPics = Pictures(sheetExport);
                string blueFile = Path.Combine(imageDir, blueId + ".png");
                string redFile = Path.Combine(imageDir, redId + ".png");
                Ok(run, "xie-sheet",
                    sheetExport.Success
                        && WdSession.Str(sheetExport, "message") == "抽取成功"
                        && sheetPics.Count == 1
                        && sheetPics[0].Id == blueId
                        && sheetPics[0].Sheet == "分析"
                        && IsExportPath(sheetPics[0].Path, stem, suffix, blueId)
                        && File.Exists(blueFile)
                        && !File.Exists(redFile)
                        && File.ReadAllText(note) == "xie-note",
                    sheetExport.Success ? "path=" + (sheetPics.Count > 0 ? sheetPics[0].Path : "") : (sheetExport.Error ?? ""));

                Console.WriteLine("只导出一张 " + redId);
                ToolResult idExport = Image(channel, "export", null, null, null, redId, null);
                List<Pic> idPics = Pictures(idExport);
                Ok(run, "xie-id",
                    idExport.Success
                        && idPics.Count == 1
                        && idPics[0].Id == redId
                        && IsExportPath(idPics[0].Path, stem, suffix, redId)
                        && File.Exists(redFile)
                        && File.Exists(blueFile)
                        && File.Exists(decoy),
                    idExport.Success ? "path=" + (idPics.Count > 0 ? idPics[0].Path : "") : (idExport.Error ?? ""));

                Console.WriteLine("导出整簿");
                ToolResult allExport = Image(channel, "export", null, null, null, null, null);
                List<Pic> allPics = Pictures(allExport);
                Ok(run, "xie-msg",
                    allExport.Success
                        && WdSession.Str(allExport, "message") == "抽取成功"
                        && WdSession.IntOf(allExport, "images_count") == 2
                        && allPics.Count == 2,
                    allExport.Success ? "count=" + WdSession.IntOf(allExport, "images_count") : (allExport.Error ?? ""));
                Ok(run, "xie-abs",
                    allExport.Success && AllExportPaths(allPics, stem, suffix),
                    PathsOf(allPics));
                Ok(run, "xie-png",
                    allExport.Success && AllPng(allPics),
                    PathsOf(allPics));
                Ok(run, "xie-dir",
                    allExport.Success
                        && AllUnder(allPics, stem, suffix)
                        && !File.Exists(Path.Combine(curr, stem, redId + ".png"))
                        && !File.Exists(Path.Combine(curr, stem + "_word", redId + ".png"))
                        && !File.Exists(Path.Combine(Path.GetDirectoryName(path) ?? "", redId + ".png")),
                    PathsOf(allPics));
                Ok(run, "xie-no-session",
                    allExport.Success
                        && !TreeHasFile(sessions, redId + ".png")
                        && !TreeHasFile(sessions, blueId + ".png"),
                    "sessions 里出现了导出 png");
                Ok(run, "xie-noclear",
                    allExport.Success
                        && File.Exists(note) && File.ReadAllText(note) == "xie-note"
                        && File.Exists(decoy) && IsPng(decoy)
                        && File.Exists(wordNote) && File.ReadAllText(wordNote) == "xie-word"
                        && File.Exists(pptNote) && File.ReadAllText(pptNote) == "xie-ppt"
                        && File.Exists(redFile) && File.Exists(blueFile),
                    "note=" + File.Exists(note) + " decoy=" + File.Exists(decoy));

                string fullAfter;
                bool savedAfter;
                bool sawAfter = TryReadBook(et, path, out _, out fullAfter, out savedAfter);
                Ok(run, "xie-book",
                    sawBefore && sawAfter && SamePath(fullBefore, fullAfter) && savedBefore == false && savedAfter == false,
                    "FullName " + fullBefore + " → " + fullAfter + " Saved " + savedBefore + " → " + savedAfter);
                string opsAfter = CkptCaseUtil.DataText(session.ListOperations(channel));
                Ok(run, "xie-ckpt",
                    opsBefore == opsAfter && opsBefore.Length > 0,
                    "导出前后检查点列表变了");

                Console.WriteLine("删掉一张后再导出");
                long blueLen = File.Exists(blueFile) ? new FileInfo(blueFile).Length : 0;
                ToolResult deleted = Image(channel, "delete", null, null, null, blueId, null);
                ToolResult afterDelete = deleted.Success
                    ? Image(channel, "export", null, null, null, null, null)
                    : deleted;
                List<Pic> left = Pictures(afterDelete);
                Ok(run, "xie-keep-gone",
                    deleted.Success
                        && afterDelete.Success
                        && WdSession.Str(afterDelete, "message") == "抽取成功"
                        && left.Count == 1
                        && left[0].Id == redId
                        && File.Exists(blueFile)
                        && new FileInfo(blueFile).Length == blueLen,
                    deleted.Success ? (afterDelete.Error ?? PathsOf(left)) : (deleted.Error ?? ""));

                string redPath = left.Count == 1 ? left[0].Path : "";
                ToolResult updated = Image(channel, "update", null, null, redPath, redId, null);
                Ok(run, "xie-reuse",
                    updated.Success
                        && WdSession.Str(updated, "image_id") == redId
                        && File.Exists(ToLocal(redPath)),
                    updated.Success ? "image_id=" + WdSession.Str(updated, "image_id") : (updated.Error ?? ""));

                string wsName = "xie-ws-" + stamp + ".png";
                string wsSession = Path.Combine(sessionDir, wsName);
                File.Copy(png, wsSession, true);
                long wsLen = new FileInfo(wsSession).Length;
                ToolResult wsIns = Image(channel, "insert", firstSheet, "D5", "workspace:" + wsName, null, null);
                string wsId = WdSession.Str(wsIns, "image_id");
                ToolResult wsExport = wsIns.Success
                    ? Image(channel, "export", null, null, null, wsId, null)
                    : wsIns;
                List<Pic> wsPics = Pictures(wsExport);
                string wsExported = wsPics.Count == 1 ? ToLocal(wsPics[0].Path) : "";
                Ok(run, "xie-ws",
                    wsIns.Success
                        && ImageIdRx.IsMatch(wsId)
                        && wsExport.Success
                        && wsPics.Count == 1
                        && IsExportPath(wsPics[0].Path, stem, suffix, wsId)
                        && File.Exists(wsSession)
                        && new FileInfo(wsSession).Length == wsLen
                        && File.Exists(wsExported)
                        && !SamePath(wsSession, wsExported),
                    wsIns.Success ? (wsExport.Error ?? wsPics[0].Path) : (wsIns.Error ?? ""));
            }
            finally
            {
                session.CloseByPath(path, channel);
            }
        }

        private static void RunEmpty(
            TestRun run,
            OfficeCkptSession session,
            string curr,
            string stamp,
            string suffix)
        {
            if (Want("xie-bad-fmt"))
            {
                ToolResult badFmt = Image(null, "export", null, null, null, "I_ab", null);
                Ok(run, "xie-bad-fmt",
                    !badFmt.Success
                        && CkptCaseUtil.Contains(badFmt.Error, "I_")
                        && !CkptCaseUtil.Contains(badFmt.Error, "抽取成功"),
                    badFmt.Success ? "短码应失败" : (badFmt.Error ?? ""));
            }

            if (!WantsAny(EmptyIds))
            {
                return;
            }

            string path = Path.Combine(Path.GetTempPath(), "ew-xie-empty-" + stamp + ".xlsx");
            string channel = null;
            try
            {
                Console.WriteLine("打开空簿 " + path);
                ToolResult opened = session.Open(path);
                if (!opened.Success)
                {
                    FailIds(run, Wanted(EmptyIds), "开簿失败: " + (opened.Error ?? ""));
                    return;
                }

                channel = WdSession.Str(opened, "channel_id");
                ToolResult overview = session.Overview(channel);
                string bookName = WdSession.Str(overview, "name");
                if (!overview.Success || string.IsNullOrEmpty(bookName))
                {
                    FailIds(run, Wanted(EmptyIds), "读不到簿名");
                    return;
                }

                string stem = CurrDocsImagePaths.SanitizeStem(bookName);
                string imageDir = Path.Combine(curr, stem + "_" + suffix);
                if (Directory.Exists(imageDir))
                {
                    FailIds(run, Wanted(EmptyIds), "空簿目录事先已存在: " + imageDir);
                    return;
                }

                Console.WriteLine("空簿导出");
                ToolResult empty = Image(channel, "export", null, null, null, null, null);
                Ok(run, "xie-empty",
                    empty.Success
                        && WdSession.Str(empty, "message") == "抽取成功"
                        && WdSession.IntOf(empty, "images_count") == 0
                        && Pictures(empty).Count == 0
                        && !Directory.Exists(imageDir),
                    empty.Success ? "目录=" + imageDir : (empty.Error ?? ""));

                ToolResult badSheet = Image(channel, "export", "没有这张表", null, null, null, null);
                Ok(run, "xie-bad-sheet",
                    FailedWithoutExtract(badSheet, "未找到工作表") && !Directory.Exists(imageDir),
                    badSheet.Success ? "未知表应失败" : (badSheet.Error ?? ""));

                ToolResult badId = Image(channel, "export", null, null, null, "I_zzzzzzzz", null);
                Ok(run, "xie-bad-id",
                    FailedWithoutExtract(badId, "未找到图片") && !Directory.Exists(imageDir),
                    badId.Success ? "未知码应失败" : (badId.Error ?? ""));
            }
            finally
            {
                session.CloseByPath(path, channel);
            }
        }

        private static void RunWord(TestRun run, string stamp)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-xie-word-" + stamp + ".docx");
            string channel = null;
            try
            {
                Console.WriteLine("打开文字稿看渠道挡板");
                ToolResult opened = OpenWordWhenFree(path);
                if (opened == null || !opened.Success)
                {
                    Fail(run, "xie-word-ch", "开文字稿失败: " + (opened == null ? "" : (opened.Error ?? "")));
                    return;
                }

                channel = WdSession.Str(opened, "channel_id");
                ToolResult exported = Image(channel, "export", null, null, null, null, null);
                Ok(run, "xie-word-ch",
                    FailedWithoutExtract(exported, "excel/et"),
                    exported.Success ? "文字渠道应失败" : (exported.Error ?? ""));
            }
            finally
            {
                if (!string.IsNullOrEmpty(channel))
                {
                    DesktopToolPipe.Call("F_close_document", new Dictionary<string, object>
                    {
                        ["channel_id"] = channel
                    });
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
        }

        private static void RunChart(
            TestRun run,
            OfficeCkptSession session,
            bool et,
            string curr,
            string stamp,
            string suffix)
        {
            string png = WritePng(stamp + "c");
            string path = Path.Combine(Path.GetTempPath(), "ew-xie-chart-" + stamp + ".xlsx");
            string channel = null;
            try
            {
                Console.WriteLine("打开带活图的簿");
                ToolResult opened = session.Open(path);
                if (!opened.Success)
                {
                    Fail(run, "xie-chart", "开簿失败: " + (opened.Error ?? ""));
                    return;
                }

                channel = WdSession.Str(opened, "channel_id");
                ToolResult overview = session.Overview(channel);
                string bookName = WdSession.Str(overview, "name");
                string sheet = FirstSheet(overview);
                if (!overview.Success || string.IsNullOrEmpty(bookName) || string.IsNullOrEmpty(sheet))
                {
                    Fail(run, "xie-chart", "读不到簿名或表名");
                    return;
                }

                string csvPath = Path.Combine(Path.GetTempPath(), "ew-xie-chart-" + stamp + ".csv");
                File.WriteAllText(
                    csvPath,
                    "年份,销量\r\n2024,1\r\n2025,2\r\n2026,3\r\n",
                    new UTF8Encoding(false));
                ToolResult written = DesktopToolPipe.Call("F_write_excel_range", new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["sheet"] = sheet,
                    ["range"] = "A1",
                    ["path"] = csvPath
                });
                string chartError = null;
                bool chartOk = written.Success && TryAddChart(et, path, sheet, out chartError);
                ToolResult ins = chartOk
                    ? Image(channel, "insert", sheet, "B2", png, null, null)
                    : written;
                string imageId = WdSession.Str(ins, "image_id");
                if (!written.Success || !chartOk || !ins.Success || !ImageIdRx.IsMatch(imageId))
                {
                    Fail(run, "xie-chart", written.Error ?? chartError ?? ins.Error ?? "活图或贴图失败");
                    return;
                }

                string stem = CurrDocsImagePaths.SanitizeStem(bookName);
                string imageDir = Path.Combine(curr, stem + "_" + suffix);
                ToolResult exported = Image(channel, "export", null, null, null, null, null);
                List<Pic> pics = Pictures(exported);
                int pngCount = Directory.Exists(imageDir)
                    ? Directory.GetFiles(imageDir, "*.png").Length
                    : 0;
                Ok(run, "xie-chart",
                    exported.Success
                        && WdSession.Str(exported, "message") == "抽取成功"
                        && pics.Count == 1
                        && pics[0].Id == imageId
                        && pngCount == 1
                        && File.Exists(Path.Combine(imageDir, imageId + ".png")),
                    exported.Success ? "png=" + pngCount + " " + PathsOf(pics) : (exported.Error ?? ""));
            }
            finally
            {
                session.CloseByPath(path, channel);
            }
        }

        private static ToolResult Image(
            string channelId,
            string action,
            string sheet,
            string anchor,
            string path,
            string imageId,
            double? widthCm)
        {
            var args = new Dictionary<string, object> { ["action"] = action };
            if (!string.IsNullOrEmpty(channelId))
            {
                args["channel_id"] = channelId;
            }

            if (sheet != null)
            {
                args["sheet"] = sheet;
            }

            if (anchor != null)
            {
                args["anchor"] = anchor;
            }

            if (path != null)
            {
                args["path"] = path;
            }

            if (imageId != null)
            {
                args["image_id"] = imageId;
            }

            if (widthCm.HasValue)
            {
                args["width_cm"] = widthCm.Value;
            }

            return DesktopToolPipe.Call("F_excel_image", args, 180000);
        }

        private static ToolResult OpenWordWhenFree(string path)
        {
            ToolResult opened = null;
            for (int attempt = 0; attempt < 6; attempt++)
            {
                opened = DesktopToolPipe.Call("F_open_document", new Dictionary<string, object>
                {
                    ["path"] = path,
                    ["create_blank"] = true,
                    ["app"] = "word"
                });
                if (opened != null && opened.Success)
                {
                    return opened;
                }

                string openError = opened == null ? "" : (opened.Error ?? "");
                bool busy = openError.IndexOf("RPC_E_CALL_REJECTED", StringComparison.OrdinalIgnoreCase) >= 0
                    || openError.IndexOf("0x80010001", StringComparison.OrdinalIgnoreCase) >= 0
                    || openError.IndexOf("拒绝接收呼叫", StringComparison.Ordinal) >= 0;
                if (!busy || attempt == 5)
                {
                    return opened;
                }

                Console.WriteLine("文字稿正忙，重试开稿");
                System.Threading.Thread.Sleep(3000);
            }

            return opened;
        }

        private static bool TryAddChart(bool et, string path, string sheetName, out string error)
        {
            error = null;
            try
            {
                object app;
                string resolveError;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out resolveError, false, false)
                    : ExcelApplicationResolver.TryResolve(out app, out resolveError, false, false);
                if (!ok || app == null)
                {
                    error = resolveError ?? "找不到表格程序";
                    return false;
                }

                dynamic application = app;
                dynamic books = application.Workbooks;
                int count = Convert.ToInt32(books.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic book = books.Item(i);
                    if (!SamePath(Convert.ToString(book.FullName), path))
                    {
                        continue;
                    }

                    dynamic sheet = book.Worksheets.Item(sheetName);
                    dynamic data = sheet.Range("A1:B4");
                    dynamic chartObj = sheet.ChartObjects().Add(220.0, 8.0, 280.0, 160.0);
                    dynamic chart = chartObj.Chart;
                    chart.ChartType = 51;
                    chart.SetSourceData(data, 2);
                    return true;
                }

                error = "找不到工作簿";
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static void DirtyCell(bool et, string path)
        {
            try
            {
                object app;
                string error;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out error, false, false)
                    : ExcelApplicationResolver.TryResolve(out app, out error, false, false);
                if (!ok || app == null)
                {
                    return;
                }

                dynamic application = app;
                dynamic books = application.Workbooks;
                int count = Convert.ToInt32(books.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic book = books.Item(i);
                    string full = Convert.ToString(book.FullName);
                    if (!SamePath(full, path))
                    {
                        continue;
                    }

                    dynamic sheet = book.ActiveSheet;
                    dynamic cell = sheet.Range("Z1");
                    cell.Value2 = "xie";
                    return;
                }
            }
            catch
            {
            }
        }

        private static bool TryReadBook(bool et, string path, out string name, out string fullName, out bool saved)
        {
            name = "";
            fullName = "";
            saved = true;
            try
            {
                object app;
                string error;
                bool ok = et
                    ? EtApplicationResolver.TryResolve(out app, out error, false, false)
                    : ExcelApplicationResolver.TryResolve(out app, out error, false, false);
                if (!ok || app == null)
                {
                    return false;
                }

                dynamic application = app;
                dynamic books = application.Workbooks;
                int count = Convert.ToInt32(books.Count);
                for (int i = 1; i <= count; i++)
                {
                    dynamic book = books.Item(i);
                    string full = Convert.ToString(book.FullName);
                    if (!SamePath(full, path))
                    {
                        continue;
                    }

                    name = Convert.ToString(book.Name);
                    fullName = full;
                    object raw = book.Saved;
                    saved = raw is bool flag ? flag : Convert.ToInt32(raw) != 0;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private sealed class Pic
        {
            public string Id;
            public string Sheet;
            public string Path;
        }

        private static List<Pic> Pictures(ToolResult result)
        {
            var list = new List<Pic>();
            Dictionary<string, object> data = WdSession.Data(result);
            if (data == null || !data.ContainsKey("images") || data["images"] == null)
            {
                return list;
            }

            JArray rows;
            try
            {
                rows = JArray.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(data["images"]));
            }
            catch
            {
                return list;
            }

            foreach (JToken row in rows)
            {
                list.Add(new Pic
                {
                    Id = row.Value<string>("image_id") ?? "",
                    Sheet = row.Value<string>("sheet") ?? "",
                    Path = row.Value<string>("path") ?? ""
                });
            }

            return list;
        }

        private static bool JsonHasPath(ToolResult result)
        {
            return CkptCaseUtil.DataText(result).IndexOf("\"path\"", StringComparison.Ordinal) >= 0;
        }

        private static bool AllExportPaths(List<Pic> pics, string stem, string suffix)
        {
            if (pics == null || pics.Count == 0)
            {
                return false;
            }

            foreach (Pic pic in pics)
            {
                if (!IsExportPath(pic.Path, stem, suffix, pic.Id))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsExportPath(string path, string stem, string suffix, string imageId)
        {
            if (string.IsNullOrEmpty(path) || path.IndexOf('\\') >= 0)
            {
                return false;
            }

            if (path.IndexOf("workspace:", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return false;
            }

            string pattern = "^[A-Za-z]:/.+/curr_docs/"
                + Regex.Escape(stem) + "_" + Regex.Escape(suffix) + "/"
                + Regex.Escape(imageId) + "\\.png$";
            return Regex.IsMatch(path, pattern);
        }

        private static bool AllUnder(List<Pic> pics, string stem, string suffix)
        {
            string needle = "/" + stem + "_" + suffix + "/";
            string bare = "/" + stem + "/";
            foreach (Pic pic in pics)
            {
                if (pic.Path.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }

                if (pic.Path.IndexOf(bare, StringComparison.OrdinalIgnoreCase) >= 0
                    && pic.Path.IndexOf(needle, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            return pics.Count > 0;
        }

        private static bool AllPng(List<Pic> pics)
        {
            if (pics == null || pics.Count == 0)
            {
                return false;
            }

            foreach (Pic pic in pics)
            {
                if (!IsPng(ToLocal(pic.Path)))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsPng(string path)
        {
            try
            {
                byte[] bytes = File.ReadAllBytes(path);
                return bytes.Length > 8
                    && bytes[0] == 0x89
                    && bytes[1] == 0x50
                    && bytes[2] == 0x4E
                    && bytes[3] == 0x47;
            }
            catch
            {
                return false;
            }
        }

        private static bool FailedWithoutExtract(ToolResult result, string needle)
        {
            return result != null
                && !result.Success
                && CkptCaseUtil.Contains(result.Error, needle)
                && !CkptCaseUtil.Contains(result.Error, "抽取成功")
                && WdSession.Str(result, "message") != "抽取成功";
        }

        private static string FirstSheet(ToolResult overview)
        {
            Dictionary<string, object> data = WdSession.Data(overview);
            if (data == null || !data.ContainsKey("sheets") || data["sheets"] == null)
            {
                return "";
            }

            try
            {
                JArray rows = JArray.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(data["sheets"]));
                if (rows.Count == 0)
                {
                    return "";
                }

                return rows[0].Value<string>("name") ?? "";
            }
            catch
            {
                return "";
            }
        }

        private static string PathsOf(List<Pic> pics)
        {
            if (pics == null || pics.Count == 0)
            {
                return "";
            }

            var parts = new List<string>();
            foreach (Pic pic in pics)
            {
                parts.Add(pic.Path);
            }

            return string.Join(" | ", parts);
        }

        private static string WritePng(string stamp)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-xie-" + stamp + ".png");
            File.WriteAllBytes(path, Convert.FromBase64String(TinyPng));
            return path;
        }

        private static string ToLocal(string path)
        {
            return string.IsNullOrEmpty(path) ? "" : path.Replace('/', '\\');
        }

        private static bool SamePath(string left, string right)
        {
            if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right))
            {
                return false;
            }

            return string.Equals(
                left.Replace('/', '\\').TrimEnd('\\'),
                right.Replace('/', '\\').TrimEnd('\\'),
                StringComparison.OrdinalIgnoreCase);
        }

        private static bool TreeHasFile(string root, string fileName)
        {
            if (string.IsNullOrEmpty(fileName) || !Directory.Exists(root))
            {
                return false;
            }

            try
            {
                foreach (string file in Directory.EnumerateFiles(root, fileName, SearchOption.AllDirectories))
                {
                    if (!string.IsNullOrEmpty(file))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryWorkspace(out string root, out string sessionDir, out string error)
        {
            root = null;
            sessionDir = null;
            error = null;
            ToolResult info = DesktopToolPipe.Call("F_workspace_info", new Dictionary<string, object>());
            root = WdSession.Str(info, "workspace_root");
            sessionDir = WdSession.Str(info, "session_dir");
            if (!info.Success || string.IsNullOrWhiteSpace(root) || string.IsNullOrWhiteSpace(sessionDir))
            {
                error = info.Error ?? "工作区根不可用";
                return false;
            }

            return true;
        }

        private static void Sweep(OfficeCkptSession session)
        {
            try
            {
                foreach (string file in Directory.GetFiles(Path.GetTempPath(), "ew-xie-*.xlsx"))
                {
                    session.CloseByPath(file);
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

        private static void AddWanted(List<string> wanted, string[] ids)
        {
            foreach (string id in ids)
            {
                if (Want(id))
                {
                    wanted.Add(id);
                }
            }
        }

        private static List<string> Wanted(string[] ids)
        {
            var list = new List<string>();
            foreach (string id in ids)
            {
                if (Want(id))
                {
                    list.Add(id);
                }
            }

            return list;
        }

        private static bool WantsAny(string[] ids)
        {
            foreach (string id in ids)
            {
                if (Want(id))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool Want(string id)
        {
            if (string.IsNullOrEmpty(IdFilter)
                || string.Equals(IdFilter, "xie", StringComparison.OrdinalIgnoreCase)
                || string.Equals(IdFilter, "xie-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return id.StartsWith(IdFilter, StringComparison.OrdinalIgnoreCase)
                || string.Equals(IdFilter, id, StringComparison.OrdinalIgnoreCase);
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            if (!Want(id))
            {
                return;
            }

            if (cond)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static void Fail(TestRun run, string id, string reason)
        {
            if (Want(id))
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static void FailIds(TestRun run, List<string> ids, string reason)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason ?? "");
            }
        }
    }
}
