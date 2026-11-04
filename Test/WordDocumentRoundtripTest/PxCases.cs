using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 读页图片落到 curr_docs/{文稿名}_ppt。前缀 px-。正式 19。环境条不在这里。
    /// </summary>
    internal static class PxCases
    {
        public const int Target = 19;

        public static string IdFilter;

        private static readonly string[] MainIds =
        {
            "px-abs-src", "px-abs-png", "px-dir", "px-no-session", "px-noclear",
            "px-same", "px-diff", "px-empty", "px-deck",
            "px-list", "px-html",
            "px-apply-abs", "px-apply-ws", "px-reread",
            "px-word-keep"
        };

        private static readonly string[] IsolatedIds =
        {
            "px-skel", "px-no-slide", "px-bad-slide", "px-word-ch"
        };

        private static readonly Regex SrcRe = new Regex(
            "data-src=\"([^\"]*)\"",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void Run(TestRun run, string app)
        {
            var wanted = new List<string>();
            foreach (string id in MainIds)
            {
                if (Want(id))
                {
                    wanted.Add(id);
                }
            }

            foreach (string id in IsolatedIds)
            {
                if (Want(id))
                {
                    wanted.Add(id);
                }
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
            Console.WriteLine("工作区 " + root);
            if (WantsAny(MainIds))
            {
                RunMain(run, app, curr, sessions, sessionDir, stamp);
            }

            if (Want("px-skel"))
            {
                RunSkel(run, app, curr, sessionDir, stamp);
            }

            if (Want("px-no-slide"))
            {
                RunNoSlide(run, app, curr, sessionDir, stamp);
            }

            if (Want("px-bad-slide"))
            {
                RunBadSlide(run, app, curr, sessionDir, stamp);
            }

            if (Want("px-word-ch"))
            {
                RunWordChannel(run, app, curr, sessionDir, stamp);
            }
        }

        private static void RunMain(
            TestRun run,
            string app,
            string curr,
            string sessions,
            string sessionDir,
            string stamp)
        {
            string red = WritePng(stamp, 220, 30, 40);
            string blue = WritePng(stamp + "b", 30, 60, 210);
            string wsPng = WritePng(stamp + "w", 20, 160, 90);
            string path = TempPptx("pxmain" + stamp);
            string openedPath = null;
            try
            {
                Console.WriteLine("打开演示文稿 " + path);
                ToolResult opened = OpenDeck(path, app);
                if (!opened.Success)
                {
                    FailIds(run, Wanted(MainIds), "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                openedPath = path;
                string channel = WdSession.Str(opened, "channel_id");
                string presName;
                string fullAtOpen;
                if (!TryReadDeck(app, path, out presName, out fullAtOpen, out _))
                {
                    FailIds(run, Wanted(MainIds), "读不到演示文稿 Name / FullName");
                    return;
                }

                string stem = CurrDocsImagePaths.SanitizeStem(presName);
                string imageDir = Path.Combine(curr, stem + "_ppt");
                Console.WriteLine("文稿 " + presName + "，补到 4 页");
                if (!EnsureSlides(channel, 4, out List<string> slides, out string slideError))
                {
                    FailIds(run, Wanted(MainIds), slideError);
                    return;
                }

                Console.WriteLine("写回第 1 页图片");
                bool picOk = ApplyPicture(channel, sessionDir, slides[0], red, stamp + "-a", out string applyError);
                if (picOk)
                {
                    Console.WriteLine("写回第 2 页图片");
                    picOk = ApplyPicture(channel, sessionDir, slides[1], red, stamp + "-b", out applyError);
                }

                if (picOk)
                {
                    Console.WriteLine("写回第 3 页图片");
                    picOk = ApplyPicture(channel, sessionDir, slides[2], blue, stamp + "-c", out applyError);
                }

                if (!picOk)
                {
                    FailIds(run, Wanted(MainIds), applyError);
                    return;
                }

                Directory.CreateDirectory(imageDir);
                string note = Path.Combine(imageDir, "note.txt");
                File.WriteAllText(note, "px-note");
                string fullBefore;
                bool savedBefore;
                bool sawBefore = TryReadDeck(app, path, out _, out fullBefore, out savedBefore);
                if (sawBefore && savedBefore)
                {
                    ApplyText(channel, sessionDir, slides[3], "px-dirty", stamp);
                    sawBefore = TryReadDeck(app, path, out _, out fullBefore, out savedBefore);
                }

                HashSet<string> pngBefore = PngNames(imageDir);
                Console.WriteLine("读第 1 页");
                ToolResult read1 = ReadSlide(channel, slides[0], "px-" + stamp + "-s1.html", true);
                string src1 = PictureSrc(read1, sessionDir);
                string fullAfter;
                bool savedAfter;
                bool sawAfter = TryReadDeck(app, path, out _, out fullAfter, out savedAfter);
                string srcFile = ToLocal(src1);

                Ok(run, "px-abs-src",
                    read1.Success && IsHashSrc(src1, stem) && src1.IndexOf("workspace:", StringComparison.OrdinalIgnoreCase) < 0,
                    read1.Success ? "data-src=" + src1 : (read1.Error ?? ""));
                Ok(run, "px-abs-png",
                    read1.Success && File.Exists(srcFile) && IsPng(srcFile) && !Path.GetFileName(srcFile).StartsWith("image_", StringComparison.OrdinalIgnoreCase),
                    read1.Success ? "文件=" + srcFile : (read1.Error ?? ""));
                Ok(run, "px-dir",
                    read1.Success
                        && src1.IndexOf("/" + stem + "_ppt/", StringComparison.OrdinalIgnoreCase) >= 0
                        && src1.IndexOf("/" + stem + "/", StringComparison.OrdinalIgnoreCase) < 0
                        && src1.IndexOf("/" + stem + "_word/", StringComparison.OrdinalIgnoreCase) < 0
                        && !File.Exists(Path.Combine(curr, stem, Path.GetFileName(srcFile)))
                        && !File.Exists(Path.Combine(curr, stem + "_word", Path.GetFileName(srcFile))),
                    "src=" + src1);
                Ok(run, "px-no-session",
                    read1.Success && !TreeHasFile(sessions, Path.GetFileName(srcFile)),
                    "sessions 里出现了 " + Path.GetFileName(srcFile));
                Ok(run, "px-noclear",
                    read1.Success && File.Exists(note) && File.ReadAllText(note) == "px-note" && File.Exists(srcFile),
                    "note=" + File.Exists(note) + " png=" + srcFile + " 目录=" + imageDir);
                Ok(run, "px-deck",
                    sawBefore && sawAfter && SamePath(fullBefore, fullAfter) && savedBefore == false && savedAfter == false,
                    "FullName " + fullBefore + " → " + fullAfter + " Saved " + savedBefore + " → " + savedAfter);

                List<string> exported = StringList(read1, "exported_images");
                string assetsDir = WdSession.Str(read1, "assets_local_dir");
                Ok(run, "px-list",
                    read1.Success
                        && exported.Count > 0
                        && AllAbsoluteExisting(exported)
                        && SamePath(assetsDir, imageDir),
                    "exported=" + string.Join(" | ", exported) + " assets=" + assetsDir);
                string htmlPath = HtmlFile(sessionDir, read1);
                string html = File.Exists(htmlPath) ? File.ReadAllText(htmlPath) : "";
                Ok(run, "px-html",
                    read1.Success
                        && !string.IsNullOrEmpty(src1)
                        && html.IndexOf(src1, StringComparison.Ordinal) >= 0
                        && exported.Exists(one => string.Equals(one, src1, StringComparison.OrdinalIgnoreCase)),
                    File.Exists(htmlPath) ? "HTML 与 exported_images 不一致。src=" + src1 : "找不到写出的 HTML " + WdSession.Str(read1, "path"));

                Console.WriteLine("读第 2 页");
                ToolResult read2 = ReadSlide(channel, slides[1], "px-" + stamp + "-s2.html", true);
                string src2 = PictureSrc(read2, sessionDir);
                int sameCount = CountFile(imageDir, Path.GetFileName(srcFile));
                Ok(run, "px-same",
                    read2.Success
                        && string.Equals(src1, src2, StringComparison.OrdinalIgnoreCase)
                        && sameCount == 1,
                    "src1=" + src1 + " src2=" + src2 + " 份数=" + sameCount);

                HashSet<string> beforeDiff = PngNames(imageDir);
                Console.WriteLine("读第 3 页");
                ToolResult read3 = ReadSlide(channel, slides[2], "px-" + stamp + "-s3.html", true);
                string src3 = PictureSrc(read3, sessionDir);
                HashSet<string> afterDiff = PngNames(imageDir);
                Ok(run, "px-diff",
                    read3.Success
                        && IsHashSrc(src3, stem)
                        && !string.Equals(src1, src3, StringComparison.OrdinalIgnoreCase)
                        && beforeDiff.Contains(Path.GetFileName(srcFile))
                        && afterDiff.Contains(Path.GetFileName(srcFile))
                        && afterDiff.Contains(Path.GetFileName(ToLocal(src3)))
                        && afterDiff.Count == beforeDiff.Count + 1,
                    "src3=" + src3 + " png " + beforeDiff.Count + " → " + afterDiff.Count);

                HashSet<string> beforeEmpty = PngNames(imageDir);
                Console.WriteLine("读无图页");
                ToolResult read4 = ReadSlide(channel, slides[3], "px-" + stamp + "-s4.html", true);
                HashSet<string> afterEmpty = PngNames(imageDir);
                Ok(run, "px-empty",
                    read4.Success && SameSet(beforeEmpty, afterEmpty),
                    read4.Success
                        ? "无图页写出了新 png " + string.Join(",", afterEmpty)
                        : (read4.Error ?? ""));

                ToolResult applied = ApplyPicture(channel, sessionDir, slides[0], src1, stamp + "-abs", out string absError)
                    ? new ToolResult { Success = true }
                    : new ToolResult { Success = false, Error = absError };
                Ok(run, "px-apply-abs",
                    applied.Success && File.Exists(srcFile),
                    applied.Success ? "绝对路径文件不见了 " + srcFile : (applied.Error ?? ""));

                string wsName = "px-ws-" + stamp + ".png";
                string wsSession = Path.Combine(sessionDir, wsName);
                File.Copy(wsPng, wsSession, true);
                long wsLen = new FileInfo(wsSession).Length;
                string wsRel = "workspace:" + wsName;
                bool wsOk = ApplyPicture(channel, sessionDir, slides[3], wsRel, stamp + "-ws", out string wsError);
                Ok(run, "px-apply-ws",
                    wsOk && File.Exists(wsSession) && new FileInfo(wsSession).Length == wsLen,
                    wsOk ? "会话文件不见了 " + wsSession : wsError);

                HashSet<string> beforeReread = PngNames(imageDir);
                Console.WriteLine("再读第 4 页");
                ToolResult reread = ReadSlide(channel, slides[3], "px-" + stamp + "-s4b.html", true);
                string rereadSrc = PictureSrc(reread, sessionDir);
                HashSet<string> afterReread = PngNames(imageDir);
                Ok(run, "px-reread",
                    reread.Success
                        && IsHashSrc(rereadSrc, stem)
                        && !beforeReread.Contains(Path.GetFileName(ToLocal(rereadSrc)))
                        && afterReread.Contains(Path.GetFileName(ToLocal(rereadSrc)))
                        && File.Exists(wsSession),
                    reread.Success ? "data-src=" + rereadSrc : (reread.Error ?? ""));

                CloseDeck(app, path);
                openedPath = null;
                RunWordKeep(run, app, curr, stem, imageDir, note, srcFile, stamp);
            }
            finally
            {
                CloseDeck(app, openedPath);
                TryDelete(red);
                TryDelete(blue);
                TryDelete(wsPng);
                TryDelete(path);
            }
        }

        private static void RunWordKeep(
            TestRun run,
            string app,
            string curr,
            string stem,
            string imageDir,
            string note,
            string srcFile,
            string stamp)
        {
            if (!Want("px-word-keep"))
            {
                return;
            }

            string textHost = app == "wpp" ? "wps" : "word";
            var text = new WdSession(textHost);
            string docx = Path.Combine(Path.GetTempPath(), stem + ".docx");
            try
            {
                ToolResult opened = OpenTextWhenFree(text, docx);
                if (opened == null || !opened.Success)
                {
                    Fail(run, "px-word-keep", "开文字稿失败: " + (opened == null ? "" : (opened.Error ?? "")));
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string png = WritePng(stamp + "k", 240, 180, 20);
                ToolResult ins = text.InsertImage(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = png
                });
                ToolResult ext = DesktopToolPipe.Call(
                    "F_extract_document_images",
                    new Dictionary<string, object> { ["channel_id"] = id });
                string wordDir = Path.Combine(curr, stem + "_word");
                bool wordImage = false;
                if (Directory.Exists(wordDir))
                {
                    foreach (string file in Directory.GetFiles(wordDir, "image_*.*"))
                    {
                        wordImage = true;
                        break;
                    }
                }

                Ok(run, "px-word-keep",
                    ins.Success && ext.Success && wordImage && File.Exists(note) && File.Exists(srcFile),
                    "插图=" + ins.Success + " 抽图=" + (ext.Success ? "成功" : (ext.Error ?? ""))
                        + " _word图=" + wordImage + " note=" + File.Exists(note) + " png=" + File.Exists(srcFile));
                TryDelete(png);
                text.CloseDocByPath(docx);
            }
            catch (Exception ex)
            {
                Fail(run, "px-word-keep", ex.Message);
            }
        }

        private static void RunSkel(TestRun run, string app, string curr, string sessionDir, string stamp)
        {
            string png = WritePng(stamp + "s", 10, 20, 30);
            string path = TempPptx("pxskel" + stamp);
            try
            {
                ToolResult opened = OpenDeck(path, app);
                if (!opened.Success)
                {
                    Fail(run, "px-skel", "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                string channel = WdSession.Str(opened, "channel_id");
                string presName;
                if (!TryReadDeck(app, path, out presName, out _, out _))
                {
                    Fail(run, "px-skel", "读不到文稿名");
                    return;
                }

                string imageDir = Path.Combine(curr, CurrDocsImagePaths.SanitizeStem(presName) + "_ppt");
                TryDeleteDir(imageDir);
                if (!EnsureSlides(channel, 1, out List<string> slides, out string slideError)
                    || !ApplyPicture(channel, sessionDir, slides[0], png, stamp + "-skel", out slideError))
                {
                    Fail(run, "px-skel", slideError);
                    return;
                }

                TryDeleteDir(imageDir);
                ToolResult read = DesktopToolPipe.Call("F_read_ppt_html", new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["slide_id"] = slides[0],
                    ["path"] = "px-" + stamp + "-skel.html",
                    ["content_only"] = true
                });
                string display = WdSession.Str(read, "display_contents");
                Ok(run, "px-skel",
                    read.Success
                        && !Directory.Exists(imageDir)
                        && display.IndexOf("data-src=", StringComparison.OrdinalIgnoreCase) < 0,
                    read.Success ? "目录在=" + Directory.Exists(imageDir) + " display 含 data-src" : (read.Error ?? ""));
            }
            finally
            {
                CloseDeck(app, path);
                TryDelete(png);
                TryDelete(path);
            }
        }

        private static void RunNoSlide(TestRun run, string app, string curr, string sessionDir, string stamp)
        {
            string path = TempPptx("pxnoslide" + stamp);
            try
            {
                ToolResult opened = OpenDeck(path, app);
                if (!opened.Success)
                {
                    Fail(run, "px-no-slide", "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                string channel = WdSession.Str(opened, "channel_id");
                string presName;
                if (!TryReadDeck(app, path, out presName, out _, out _))
                {
                    Fail(run, "px-no-slide", "读不到文稿名");
                    return;
                }

                string imageDir = Path.Combine(curr, CurrDocsImagePaths.SanitizeStem(presName) + "_ppt");
                TryDeleteDir(imageDir);
                ToolResult read = DesktopToolPipe.Call("F_read_ppt_html", new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["path"] = "px-" + stamp + "-noslide.html"
                });
                Ok(run, "px-no-slide",
                    !read.Success && !Directory.Exists(imageDir) && string.IsNullOrEmpty(WdSession.Str(read, "display_contents")),
                    Dump(read) + " 目录在=" + Directory.Exists(imageDir));
            }
            finally
            {
                CloseDeck(app, path);
                TryDelete(path);
            }
        }

        private static void RunBadSlide(TestRun run, string app, string curr, string sessionDir, string stamp)
        {
            string png = WritePng(stamp + "bad", 80, 80, 80);
            string path = TempPptx("pxbad" + stamp);
            try
            {
                ToolResult opened = OpenDeck(path, app);
                if (!opened.Success)
                {
                    Fail(run, "px-bad-slide", "开稿失败: " + (opened.Error ?? ""));
                    return;
                }

                string channel = WdSession.Str(opened, "channel_id");
                string presName;
                if (!TryReadDeck(app, path, out presName, out _, out _))
                {
                    Fail(run, "px-bad-slide", "读不到文稿名");
                    return;
                }

                string imageDir = Path.Combine(curr, CurrDocsImagePaths.SanitizeStem(presName) + "_ppt");
                TryDeleteDir(imageDir);
                ToolResult read = DesktopToolPipe.Call("F_read_ppt_html", new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["slide_id"] = "999999999",
                    ["path"] = "px-" + stamp + "-bad.html"
                });
                Ok(run, "px-bad-slide",
                    !read.Success && PngNames(imageDir).Count == 0,
                    Dump(read) + " png=" + PngNames(imageDir).Count);
            }
            finally
            {
                CloseDeck(app, path);
                TryDelete(png);
                TryDelete(path);
            }
        }

        private static void RunWordChannel(TestRun run, string app, string curr, string sessionDir, string stamp)
        {
            string textHost = app == "wpp" ? "wps" : "word";
            var text = new WdSession(textHost);
            string docx = Path.Combine(Path.GetTempPath(), "pxword" + stamp + ".docx");
            try
            {
                ToolResult opened = OpenTextWhenFree(text, docx);
                if (opened == null || !opened.Success)
                {
                    Fail(run, "px-word-ch", "开文字稿失败: " + (opened == null ? "" : (opened.Error ?? "")));
                    return;
                }

                string id = WdSession.Str(opened, "channel_id");
                string stem = CurrDocsImagePaths.SanitizeStem(Path.GetFileName(docx));
                string imageDir = Path.Combine(curr, stem + "_ppt");
                TryDeleteDir(imageDir);
                ToolResult read = DesktopToolPipe.Call("F_read_ppt_html", new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["slide_id"] = "1",
                    ["path"] = "px-" + stamp + "-word.html"
                });
                string error = read.Error ?? "";
                Ok(run, "px-word-ch",
                    !read.Success
                        && (error.IndexOf("渠道", StringComparison.Ordinal) >= 0 || error.IndexOf("ppt", StringComparison.OrdinalIgnoreCase) >= 0)
                        && !Directory.Exists(imageDir),
                    Dump(read) + " 目录在=" + Directory.Exists(imageDir));
                text.CloseDocByPath(docx);
            }
            catch (Exception ex)
            {
                Fail(run, "px-word-ch", ex.Message);
            }
        }

        private static bool EnsureSlides(string channel, int count, out List<string> slides, out string error)
        {
            slides = new List<string>();
            error = null;
            for (int i = 0; i < 8 && slides.Count < count; i++)
            {
                ToolResult content = DesktopToolPipe.Call(
                    "F_get_presentation_content",
                    new Dictionary<string, object> { ["channel_id"] = channel });
                slides = SlideIds(content);
                if (!content.Success)
                {
                    error = content.Error ?? "读不到页列表";
                    return false;
                }

                if (slides.Count >= count)
                {
                    return true;
                }

                ToolResult added = DesktopToolPipe.Call("F_manage_ppt_slide", new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["action"] = "add"
                });
                if (!added.Success)
                {
                    error = added.Error ?? "加页失败";
                    return false;
                }
            }

            error = "页数不够，只有 " + slides.Count;
            return slides.Count >= count;
        }

        private static bool ApplyPicture(
            string channel,
            string sessionDir,
            string slideId,
            string src,
            string tag,
            out string error)
        {
            error = null;
            string rel = "px-" + tag + ".html";
            string html = "<section style=\"position:relative;width:100%;height:100%\">\n"
                + "  <img data-shape-type=\"picture\" data-src=\"" + EscapeAttr(src) + "\""
                + " style=\"position:absolute;left:8%;top:10%;width:30%;height:24%\" />\n"
                + "</section>";
            ToolResult written = DesktopToolPipe.Call("F_write_file", new Dictionary<string, object>
            {
                ["path"] = rel,
                ["content"] = html
            });
            if (!written.Success)
            {
                error = written.Error ?? "写 HTML 失败";
                return false;
            }

            ToolResult applied = DesktopToolPipe.Call("F_apply_ppt_html", new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["slide_id"] = slideId,
                ["path"] = rel,
                ["allow_create"] = true
            });
            if (!applied.Success)
            {
                error = applied.Error ?? "写回图片失败";
                return false;
            }

            return true;
        }

        private static void ApplyText(string channel, string sessionDir, string slideId, string text, string stamp)
        {
            string rel = "px-" + stamp + "-dirty.html";
            string html = "<section style=\"position:relative;width:100%;height:100%\">\n"
                + "  <div data-shape-type=\"textbox\" style=\"position:absolute;left:8%;top:40%;width:40%;height:16%\">"
                + text + "</div>\n</section>";
            DesktopToolPipe.Call("F_write_file", new Dictionary<string, object>
            {
                ["path"] = rel,
                ["content"] = html
            });
            DesktopToolPipe.Call("F_apply_ppt_html", new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["slide_id"] = slideId,
                ["path"] = rel,
                ["allow_create"] = true
            });
        }

        private static ToolResult ReadSlide(string channel, string slideId, string rel, bool full)
        {
            var args = new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["slide_id"] = slideId,
                ["path"] = rel
            };
            if (full)
            {
                args["full"] = true;
            }

            return DesktopToolPipe.Call("F_read_ppt_html", args);
        }

        /// <summary>
        /// full=true 的回包 display 是摘要，data-src 在写出的 HTML 和 exported_images 里。
        /// </summary>
        private static string PictureSrc(ToolResult read, string sessionDir)
        {
            string htmlPath = HtmlFile(sessionDir, read);
            if (File.Exists(htmlPath))
            {
                string fromHtml = FirstSrc(File.ReadAllText(htmlPath));
                if (!string.IsNullOrEmpty(fromHtml))
                {
                    return fromHtml;
                }
            }

            List<string> exported = StringList(read, "exported_images");
            return exported.Count > 0 ? exported[0] : "";
        }

        private static ToolResult OpenTextWhenFree(WdSession text, string docx)
        {
            ToolResult opened = null;
            for (int attempt = 0; attempt < 6; attempt++)
            {
                opened = text.Open(docx, true, text.App);
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

        private static ToolResult OpenDeck(string path, string app)
        {
            return DesktopToolPipe.Call("F_open_document", new Dictionary<string, object>
            {
                ["path"] = path,
                ["create_blank"] = true,
                ["app"] = app == "wpp" ? "wpp" : "powerpoint"
            });
        }

        private static bool IsHashSrc(string src, string stem)
        {
            if (string.IsNullOrEmpty(src) || src.IndexOf('\\') >= 0)
            {
                return false;
            }

            return Regex.IsMatch(
                src,
                "^[A-Za-z]:/.+/curr_docs/" + Regex.Escape(stem) + "_ppt/[0-9a-fA-F]{8}\\.png$");
        }

        private static string FirstSrc(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            Match m = SrcRe.Match(display);
            return m.Success ? m.Groups[1].Value : "";
        }

        private static List<string> SlideIds(ToolResult result)
        {
            var list = new List<Tuple<int, string>>();
            Dictionary<string, object> data = WdSession.Data(result);
            if (data != null && data.ContainsKey("slides") && data["slides"] != null)
            {
                JToken token = data["slides"] as JToken ?? JToken.FromObject(data["slides"]);
                foreach (JToken item in token)
                {
                    string id = item["slide_id"] == null ? "" : item["slide_id"].ToString();
                    int index = 0;
                    if (item["index"] != null)
                    {
                        int.TryParse(item["index"].ToString(), out index);
                    }

                    if (!string.IsNullOrEmpty(id))
                    {
                        list.Add(Tuple.Create(index, id));
                    }
                }
            }

            list.Sort((a, b) => a.Item1.CompareTo(b.Item1));
            var ids = new List<string>();
            foreach (Tuple<int, string> one in list)
            {
                ids.Add(one.Item2);
            }

            return ids;
        }

        private static List<string> StringList(ToolResult result, string key)
        {
            var list = new List<string>();
            Dictionary<string, object> data = WdSession.Data(result);
            if (data == null || !data.ContainsKey(key) || data[key] == null)
            {
                return list;
            }

            JToken token = data[key] as JToken ?? JToken.FromObject(data[key]);
            foreach (JToken item in token)
            {
                string text = item == null ? "" : item.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    list.Add(text);
                }
            }

            return list;
        }

        private static bool AllAbsoluteExisting(List<string> paths)
        {
            foreach (string one in paths)
            {
                if (!IsRootedForward(one) || !File.Exists(ToLocal(one)))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsRootedForward(string path)
        {
            return !string.IsNullOrEmpty(path)
                && path.IndexOf('\\') < 0
                && Regex.IsMatch(path, "^[A-Za-z]:/");
        }

        private static string HtmlFile(string sessionDir, ToolResult read)
        {
            string rel = WdSession.Str(read, "path");
            if (string.IsNullOrEmpty(rel))
            {
                rel = WdSession.Str(read, "html_filename");
            }

            if (string.IsNullOrEmpty(rel))
            {
                return "";
            }

            string local = rel.IndexOf(":/", StringComparison.Ordinal) >= 0 || Path.IsPathRooted(rel)
                ? ToLocal(rel)
                : Path.Combine(sessionDir, rel.Replace('/', '\\'));
            return local;
        }

        private static string ToLocal(string path)
        {
            return (path ?? "").Replace('/', '\\');
        }

        private static bool IsPng(string path)
        {
            try
            {
                using (var stream = File.OpenRead(path))
                {
                    if (stream.Length < 8)
                    {
                        return false;
                    }

                    var magic = new byte[4];
                    stream.Read(magic, 0, 4);
                    return magic[0] == 0x89 && magic[1] == 0x50 && magic[2] == 0x4E && magic[3] == 0x47;
                }
            }
            catch
            {
                return false;
            }
        }

        private static HashSet<string> PngNames(string dir)
        {
            var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (!Directory.Exists(dir))
            {
                return set;
            }

            foreach (string file in Directory.GetFiles(dir, "*.png"))
            {
                set.Add(Path.GetFileName(file));
            }

            return set;
        }

        private static int CountFile(string dir, string name)
        {
            if (!Directory.Exists(dir) || string.IsNullOrEmpty(name))
            {
                return 0;
            }

            int n = 0;
            foreach (string file in Directory.GetFiles(dir, "*.png"))
            {
                if (string.Equals(Path.GetFileName(file), name, StringComparison.OrdinalIgnoreCase))
                {
                    n++;
                }
            }

            return n;
        }

        private static bool SameSet(HashSet<string> a, HashSet<string> b)
        {
            if (a.Count != b.Count)
            {
                return false;
            }

            foreach (string one in a)
            {
                if (!b.Contains(one))
                {
                    return false;
                }
            }

            return true;
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

        private static bool TryReadDeck(string app, string path, out string name, out string fullName, out bool saved)
        {
            name = "";
            fullName = "";
            saved = true;
            if (app == "wpp")
            {
                return TryReadWpp(path, out name, out fullName, out saved);
            }

            try
            {
                var application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                string norm = Normalize(path);
                foreach (PowerPoint.Presentation pres in application.Presentations)
                {
                    string full = "";
                    try
                    {
                        full = pres.FullName ?? "";
                    }
                    catch
                    {
                    }

                    if (!SamePath(full, norm) && !SamePath(full, path))
                    {
                        continue;
                    }

                    name = pres.Name ?? "";
                    fullName = full;
                    saved = (int)pres.Saved != 0;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool TryReadWpp(string path, out string name, out string fullName, out bool saved)
        {
            name = "";
            fullName = "";
            saved = true;
            foreach (string prog in WppProgIds)
            {
                object application = null;
                try
                {
                    application = Marshal.GetActiveObject(prog);
                }
                catch
                {
                    continue;
                }

                if (TryMatchPresentation(application, path, out name, out fullName, out saved))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryMatchPresentation(
            object application,
            string path,
            out string name,
            out string fullName,
            out bool saved)
        {
            name = "";
            fullName = "";
            saved = true;
            object presentations;
            try
            {
                presentations = GetProp(application, "Presentations");
            }
            catch
            {
                return false;
            }

            int count;
            try
            {
                count = Convert.ToInt32(GetProp(presentations, "Count"));
            }
            catch
            {
                return false;
            }

            for (int i = 1; i <= count; i++)
            {
                object pres = Invoke(presentations, "Item", i);
                string full = Convert.ToString(GetProp(pres, "FullName")) ?? "";
                if (!SamePath(full, path))
                {
                    continue;
                }

                name = Convert.ToString(GetProp(pres, "Name")) ?? "";
                fullName = full;
                saved = Convert.ToInt32(GetProp(pres, "Saved")) != 0;
                return true;
            }

            return false;
        }

        private static void CloseDeck(string app, string path)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            if (app == "wpp")
            {
                CloseWpp(path);
                return;
            }

            try
            {
                var application = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application");
                var list = new List<PowerPoint.Presentation>();
                foreach (PowerPoint.Presentation pres in application.Presentations)
                {
                    list.Add(pres);
                }

                foreach (PowerPoint.Presentation pres in list)
                {
                    string full = "";
                    try
                    {
                        full = pres.FullName ?? "";
                    }
                    catch
                    {
                    }

                    if (!SamePath(full, path))
                    {
                        continue;
                    }

                    try
                    {
                        pres.Saved = Office.MsoTriState.msoTrue;
                        pres.Close();
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

        private static void CloseWpp(string path)
        {
            foreach (string prog in WppProgIds)
            {
                object application;
                try
                {
                    application = Marshal.GetActiveObject(prog);
                }
                catch
                {
                    continue;
                }

                object presentations;
                try
                {
                    presentations = GetProp(application, "Presentations");
                }
                catch
                {
                    continue;
                }

                int count;
                try
                {
                    count = Convert.ToInt32(GetProp(presentations, "Count"));
                }
                catch
                {
                    continue;
                }

                for (int i = count; i >= 1; i--)
                {
                    object pres;
                    try
                    {
                        pres = Invoke(presentations, "Item", i);
                    }
                    catch
                    {
                        continue;
                    }

                    string full = Convert.ToString(GetProp(pres, "FullName")) ?? "";
                    if (!SamePath(full, path))
                    {
                        continue;
                    }

                    try
                    {
                        SetProp(pres, "Saved", true);
                        Invoke(pres, "Close");
                    }
                    catch
                    {
                    }
                }
            }
        }

        private static readonly string[] WppProgIds =
        {
            "Kwpp.Application",
            "kwpp.Application",
            "wpp.Application"
        };

        private static object GetProp(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static void SetProp(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                BindingFlags.SetProperty,
                null,
                target,
                new[] { value });
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                BindingFlags.InvokeMethod,
                null,
                target,
                args == null || args.Length == 0 ? null : args);
        }

        private static string WritePng(string stamp, byte r, byte g, byte b)
        {
            string path = Path.Combine(Path.GetTempPath(), "px-" + stamp + ".png");
            using (var bitmap = new Bitmap(4, 4))
            {
                for (int y = 0; y < 4; y++)
                {
                    for (int x = 0; x < 4; x++)
                    {
                        bitmap.SetPixel(x, y, Color.FromArgb(255, r, g, (byte)(b + x + y)));
                    }
                }

                bitmap.Save(path, ImageFormat.Png);
            }

            return path;
        }

        private static string TempPptx(string name)
        {
            return Path.Combine(Path.GetTempPath(), name + ".pptx");
        }

        private static string EscapeAttr(string value)
        {
            return (value ?? "").Replace("&", "&amp;").Replace("\"", "&quot;");
        }

        private static bool SamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
            {
                return false;
            }

            return string.Equals(Normalize(a), Normalize(b), StringComparison.OrdinalIgnoreCase);
        }

        private static string Normalize(string path)
        {
            return (path ?? "").Replace('/', '\\').Trim();
        }

        private static bool Want(string id)
        {
            string filter = IdFilter ?? "";
            if (filter.Length == 0
                || string.Equals(filter, "px", StringComparison.OrdinalIgnoreCase)
                || string.Equals(filter, "px-", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return id.StartsWith(filter, StringComparison.OrdinalIgnoreCase);
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

        private static string Dump(ToolResult result)
        {
            if (result == null)
            {
                return "(null)";
            }

            return "success=" + result.Success + " error=" + (result.Error ?? "");
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
                run.CaseFail(id, reason);
            }
        }

        private static void Fail(TestRun run, string id, string reason)
        {
            if (Want(id))
            {
                run.CaseFail(id, reason);
            }
        }

        private static void FailIds(TestRun run, IEnumerable<string> ids, string reason)
        {
            foreach (string id in ids)
            {
                Fail(run, id, reason);
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    File.Delete(path);
                }
            }
            catch
            {
            }
        }

        private static void TryDeleteDir(string path)
        {
            try
            {
                if (!string.IsNullOrEmpty(path) && Directory.Exists(path))
                {
                    Directory.Delete(path, true);
                }
            }
            catch
            {
            }
        }
    }
}
