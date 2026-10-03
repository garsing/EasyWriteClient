using System;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>第 5 波 download，只测 agent。</summary>
    internal static class DlCases
    {
        public static void Run(TestRun run, FixtureHttp http, int chromeBefore, int edgeBefore)
        {
            var s = new NavSession();
            byte[] txtBytes = File.ReadAllBytes(http.LocalFile("download", "sample.txt"));
            byte[] pngBytes = File.ReadAllBytes(http.LocalFile("download", "sample.png"));
            RunNeg(run, s, http);
            RunUrl(run, s, http, txtBytes, pngBytes);
            RunRef(run, s, http, txtBytes);
            RunIframeRowNeg(run, s, http);
            int chromeNow = IntCases.CountProc("chrome");
            int edgeNow = IntCases.CountProc("msedge");
            if (chromeNow < chromeBefore || edgeNow < edgeBefore)
            {
                Console.WriteLine(
                    "  [WARN] 用户浏览器进程变少 chrome "
                    + chromeBefore + "→" + chromeNow
                    + " edge " + edgeBefore + "→" + edgeNow);
            }
        }

        private static void RunNeg(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-neg-neither");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 夹具",
                    "dl-neg-both", "dl-neg-no-path", "dl-neg-file-url", "dl-neg-javascript",
                    "dl-neg-bad-ref", "dl-neg-stale-ref", "dl-neg-wrong-kind", "dl-neg-unknown-id",
                    "dl-neg-image-ref", "dl-neg-plain", "dl-neg-no-event");
            }
            else
            {
                string tree = IntCases.WaitTree(s, id, false, "下载附件", "普通按钮", "示例图");
                string fileRef = FindDlRef(tree, "下载附件");
                string any = fileRef ?? "e1";
                string txtUrl = http.Url("download", "sample.txt");

                IntCases.ExpectFail(run, "dl-neg-neither", s.Download(id, path: "browser_dl/neg.txt"), "ref");
                IntCases.ExpectFail(run, "dl-neg-both", s.Download(id, refId: any, url: txtUrl, path: "browser_dl/neg.txt"), "ref");
                IntCases.ExpectFail(run, "dl-neg-no-path", s.Download(id, url: txtUrl), "path");
                IntCases.ExpectFail(run, "dl-neg-file-url", s.Download(id, url: "file:///C:/Windows/win.ini", path: "browser_dl/neg.txt"), "http");
                IntCases.ExpectFail(run, "dl-neg-javascript", s.Download(id, url: "javascript:alert(1)", path: "browser_dl/neg.txt"), "http");
                IntCases.ExpectFail(run, "dl-neg-bad-ref", s.Download(id, refId: "e9999", path: "browser_dl/neg.txt"), "过期");
                IntCases.ExpectFail(run, "dl-neg-wrong-kind", s.Download("w1", url: txtUrl, path: "browser_dl/neg.txt"), "渠道");
                IntCases.ExpectFail(run, "dl-neg-unknown-id", s.Download("b999", url: txtUrl, path: "browser_dl/neg.txt"), "渠道");
                IntCases.ExpectFail(
                    run,
                    "dl-neg-image-ref",
                    s.Download(id, refId: FindDlRef(tree, "示例图") ?? "e1", path: "browser_dl/neg.txt"),
                    "url");
                IntCases.ExpectFail(
                    run,
                    "dl-neg-plain",
                    s.Download(id, refId: FindDlRef(tree, "普通按钮") ?? "e1", path: "browser_dl/neg.txt"),
                    "下载");
                IntCases.ExpectFail(
                    run,
                    "dl-neg-no-event",
                    s.Download(id, refId: FindFakeDownload(tree) ?? "e1", path: "browser_dl/neg.txt"),
                    "下载");

                if (string.IsNullOrEmpty(fileRef))
                {
                    run.CaseFail("dl-neg-stale-ref", "空拍无下载附件");
                }
                else
                {
                    ToolResult jump = s.Nav(http.Url("public-example", "index.html"), host: "yiwrite", channelId: id);
                    if (!jump.Success)
                    {
                        run.CaseFail("dl-neg-stale-ref", jump.Error ?? "跳转失败");
                    }
                    else
                    {
                        IntCases.ExpectFail(
                            run,
                            "dl-neg-stale-ref",
                            s.Download(id, refId: fileRef, path: "browser_dl/neg.txt"),
                            "过期");
                    }
                }
            }

            string d18 = IntCases.OpenOk(run, s, http.Url("d18-risk", "index.html"), "dl-neg-password");
            if (d18 == null)
            {
                run.CaseFail("dl-neg-login", "依赖 d18-risk");
            }
            else
            {
                string t = IntCases.WaitTree(s, d18, true, "密码", "登录");
                IntCases.ExpectFail(
                    run,
                    "dl-neg-password",
                    s.Download(d18, refId: IntCases.FindRef(t, "密码") ?? "e1", path: "browser_dl/neg.txt"),
                    "密码");
                IntCases.ExpectFail(
                    run,
                    "dl-neg-login",
                    s.Download(d18, refId: IntCases.FindRef(t, "登录") ?? "e1", path: "browser_dl/neg.txt"),
                    "登录");
            }

            string chooser = IntCases.OpenOk(run, s, http.Url("file-chooser", "index.html"), "dl-neg-file-chooser");
            if (chooser == null)
            {
                return;
            }

            string ct = IntCases.WaitTree(s, chooser, false, "选择文件");
            IntCases.ExpectFail(
                run,
                "dl-neg-file-chooser",
                s.Download(chooser, refId: IntCases.FindRef(ct, "选择文件") ?? "e1", path: "browser_dl/neg.txt"),
                "文件");
        }

        private static void RunIframeRowNeg(TestRun run, NavSession s, FixtureHttp http)
        {
            string ifr = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-neg-iframe-row");
            if (ifr == null)
            {
                return;
            }

            string ft = IntCases.WaitTree(s, ifr, false, "附件框");
            string frame = IntCases.FindRef(ft, "附件框") ?? IntCases.FindIframeRef(ft);
            IntCases.ExpectFail(
                run,
                "dl-neg-iframe-row",
                s.Download(ifr, refId: string.IsNullOrEmpty(frame) ? "e1" : frame, path: "browser_dl/neg.txt"),
                "展开");
        }

        private static void RunUrl(TestRun run, NavSession s, FixtureHttp http, byte[] txtBytes, byte[] pngBytes)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-url-txt");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 夹具",
                    "dl-url-png", "dl-url-relpath", "dl-url-no-leak", "dl-url-resnapshot",
                    "dl-url-hidden", "dl-url-twice");
                return;
            }

            ToolResult txt = s.Download(id, url: http.Url("download", "sample.txt"), path: "browser_dl/sample.txt");
            bool txtOk = IsOkFile(txt, txtBytes.Length) && WorkspaceHas(s, Filename(txt), "easywrite-dl-sample");
            IntCases.Ok(run, "dl-url-txt", txtOk, Why(txt, "url txt"));
            IntCases.Ok(run, "dl-url-relpath", txt != null && txt.Success && IsRelative(Filename(txt)), Filename(txt));
            IntCases.Ok(run, "dl-url-no-leak", txt != null && txt.Success && !HasLeak(txt), "回包泄漏");
            IntCases.Ok(
                run,
                "dl-url-resnapshot",
                txt != null && txt.Success
                    && IntCases.Flag(txt, "resnapshot")
                    && !IntCases.Flag(txt, "refs_invalidated")
                    && IntCases.Tree(txt).Length > 0,
                txt == null
                    ? "无回包"
                    : ("url 未回树 resnap=" + IntCases.Flag(txt, "resnapshot")
                        + " inv=" + IntCases.Flag(txt, "refs_invalidated")
                        + " tree=" + IntCases.Tree(txt).Length
                        + " " + (txt.Error ?? "")));

            ToolResult png = s.Download(id, url: http.Url("download", "sample.png"), path: "browser_dl/sample.png");
            IntCases.Ok(
                run,
                "dl-url-png",
                IsOkFile(png, pngBytes.Length) || (png != null && png.Success && BytesOf(png) > 0),
                Why(png, "url png"));

            ToolResult twice = s.Download(id, url: http.Url("download", "sample.txt"), path: "browser_dl/sample.txt");
            IntCases.Ok(run, "dl-url-twice", twice != null && twice.Success && IsRelative(Filename(twice)), Why(twice, "再下"));

            string hid = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-url-hidden", visible: false);
            if (hid == null)
            {
                return;
            }

            ToolResult hidden = s.Download(hid, url: http.Url("download", "sample.txt"), path: "browser_dl/hidden.txt");
            IntCases.Ok(run, "dl-url-hidden", hidden != null && hidden.Success, Why(hidden, "隐藏窗 url"));
        }

        private static void RunRef(TestRun run, NavSession s, FixtureHttp http, byte[] txtBytes)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-ref-attr");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 夹具",
                    "dl-ref-export", "dl-ref-resnapshot", "dl-ref-bytes",
                    "dl-ref-hidden", "dl-ref-iframe", "dl-ref-new-tree");
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "下载附件", "导出");
            string attr = FindDlRef(tree, "下载附件");
            ToolResult got = string.IsNullOrEmpty(attr)
                ? null
                : s.Download(id, refId: attr, path: "browser_dl/ref-attr.txt");
            bool ok = IsOkFile(got, txtBytes.Length) && WorkspaceHas(s, Filename(got), "easywrite-dl-sample");
            IntCases.Ok(run, "dl-ref-attr", ok, Why(got, "ref 附件"));
            IntCases.Ok(run, "dl-ref-bytes", got != null && got.Success && BytesOf(got) == txtBytes.Length, "bytes 不对");
            IntCases.Ok(
                run,
                "dl-ref-resnapshot",
                got != null && got.Success
                    && IntCases.Flag(got, "resnapshot")
                    && !IntCases.Flag(got, "refs_invalidated")
                    && IntCases.Tree(got).Length > 0,
                "ref 未回树");

            string after = got != null && got.Success ? IntCases.Tree(got) : tree;
            string export = FindDlRef(after, "导出") ?? FindDlRef(IntCases.WaitTree(s, id, false, "导出"), "导出");
            ToolResult exp = string.IsNullOrEmpty(export)
                ? null
                : s.Download(id, refId: export, path: "browser_dl/ref-export.txt");
            IntCases.Ok(run, "dl-ref-export", exp != null && exp.Success && IsRelative(Filename(exp)), Why(exp, "导出"));

            string neu = exp != null && exp.Success ? IntCases.Tree(exp) : after;
            string again = FindDlRef(neu, "下载附件") ?? FindDlRef(IntCases.WaitTree(s, id, false, "下载附件"), "下载附件");
            ToolResult second = string.IsNullOrEmpty(again)
                ? null
                : s.Download(id, refId: again, path: "browser_dl/ref-again.txt");
            IntCases.Ok(
                run,
                "dl-ref-new-tree",
                second != null && second.Success,
                second == null ? "新树无下载附件" : second.Error);

            string hid = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-ref-hidden", visible: false);
            if (hid == null)
            {
                run.CaseFail("dl-ref-iframe", "依赖隐藏窗开页");
                return;
            }

            string ht = IntCases.WaitTree(s, hid, false, "下载附件");
            ToolResult href = s.Download(hid, refId: FindDlRef(ht, "下载附件"), path: "browser_dl/ref-hidden.txt");
            IntCases.Ok(run, "dl-ref-hidden", href != null && href.Success, Why(href, "隐藏窗 ref"));

            string ifr = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-ref-iframe");
            if (ifr == null)
            {
                return;
            }

            string ot = IntCases.WaitTree(s, ifr, false, "附件框");
            string frame = IntCases.FindRef(ot, "附件框") ?? IntCases.FindIframeRef(ot);
            ToolResult expFrame = string.IsNullOrEmpty(frame) ? null : s.Interact(ifr, "click", frame);
            if (expFrame == null || !expFrame.Success)
            {
                expFrame = string.IsNullOrEmpty(frame) ? null : s.Snap(ifr, frame);
            }

            string inner = expFrame == null ? "" : IntCases.Tree(expFrame);
            string note = FindDlRef(inner, "下载子页备注");
            ToolResult innerDl = string.IsNullOrEmpty(note)
                ? null
                : s.Download(ifr, refId: note, path: "browser_dl/ref-iframe.txt");
            IntCases.Ok(
                run,
                "dl-ref-iframe",
                innerDl != null && innerDl.Success && WorkspaceHas(s, Filename(innerDl), "easywrite-dl-sample"),
                innerDl == null ? "展开后无子页备注" : Why(innerDl, "iframe 下载"));
        }

        private static bool IsOkFile(ToolResult r, int expectBytes)
        {
            return r != null && r.Success && IsRelative(Filename(r)) && BytesOf(r) == expectBytes;
        }

        private static bool WorkspaceHas(NavSession s, string path, string hint)
        {
            if (string.IsNullOrEmpty(path))
            {
                return false;
            }

            ToolResult read = s.ReadFile(path);
            if (read == null || !read.Success)
            {
                return false;
            }

            string text = IntCases.Str(read, "content") + IntCases.Tree(read) + (read.Error ?? "");
            return text.IndexOf(hint, StringComparison.Ordinal) >= 0;
        }

        private static bool IsRelative(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            if (path.IndexOf(":\\", StringComparison.Ordinal) >= 0
                || path.StartsWith("\\\\", StringComparison.Ordinal))
            {
                return false;
            }

            return path.IndexOf("browser_dl/", StringComparison.OrdinalIgnoreCase) >= 0
                || path.IndexOf('/') >= 0
                || path.IndexOf('.') >= 0;
        }

        private static bool HasLeak(ToolResult r)
        {
            if (r == null)
            {
                return true;
            }

            var data = IntCases.Data(r);
            if (data.ContainsKey("content") || data.ContainsKey("base64") || data.ContainsKey("image_base64"))
            {
                return true;
            }

            string dump = Filename(r);
            return dump.IndexOf(":\\", StringComparison.Ordinal) >= 0;
        }

        private static string Filename(ToolResult r)
        {
            string a = IntCases.Str(r, "filename");
            if (!string.IsNullOrEmpty(a))
            {
                return a;
            }

            return IntCases.Str(r, "path");
        }

        private static int BytesOf(ToolResult r)
        {
            object v;
            if (!IntCases.Data(r).TryGetValue("bytes", out v) || v == null)
            {
                return -1;
            }

            try
            {
                return Convert.ToInt32(v);
            }
            catch
            {
                return -1;
            }
        }

        private static string Why(ToolResult r, string what)
        {
            if (r == null)
            {
                return "无回包 " + what;
            }

            if (!r.Success)
            {
                return r.Error;
            }

            return what + " filename=" + Filename(r) + " bytes=" + BytesOf(r);
        }

        private static string FindDlRef(string tree, string label)
        {
            return IntCases.FindRef(tree, "\"" + label + "\"")
                ?? IntCases.FindRef(tree, label);
        }

        private static string FindFakeDownload(string tree)
        {
            if (string.IsNullOrEmpty(tree))
            {
                return null;
            }

            foreach (string line in tree.Split('\n'))
            {
                if (line.IndexOf("下载", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (line.IndexOf("附件", StringComparison.Ordinal) >= 0
                    || line.IndexOf("子页", StringComparison.Ordinal) >= 0
                    || line.IndexOf("导出", StringComparison.Ordinal) >= 0
                    || line.IndexOf("夹具", StringComparison.Ordinal) >= 0)
                {
                    continue;
                }

                Match m = Regex.Match(line, @"ref=(e\d+)");
                if (!m.Success)
                {
                    continue;
                }

                if (line.IndexOf("button", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return m.Groups[1].Value;
                }
            }

            return null;
        }
    }
}
