using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>第 5 波 download 巩固 50 条，只测 agent。</summary>
    internal static partial class DlCases
    {
        private static readonly string[] ExtraIds =
        {
            "dl-x-neg-mailto", "dl-x-neg-data", "dl-x-neg-ftp", "dl-x-neg-about", "dl-x-neg-js-void",
            "dl-x-neg-relative", "dl-x-neg-e0", "dl-x-neg-e00", "dl-x-neg-both-png",
            "dl-x-neg-wrong-kind-ref", "dl-x-neg-unknown-ref", "dl-x-neg-plain2", "dl-x-neg-image2",
            "dl-x-neg-iframe2", "dl-x-neg-stale2", "dl-x-neg-file2", "dl-x-neg-neither2", "dl-x-neg-nopath2",
            "dl-x-neg-pay", "dl-x-neg-delete", "dl-x-neg-submit", "dl-x-neg-pay2", "dl-x-neg-signin", "dl-x-neg-delen",
            "dl-x-url-txt", "dl-x-url-txt-bytes", "dl-x-url-txt-rel", "dl-x-url-txt-noleak", "dl-x-url-txt-resnap",
            "dl-x-url-txt-source", "dl-x-url-txt-kind", "dl-x-url-txt-track",
            "dl-x-url-png", "dl-x-url-png-rel", "dl-x-url-png-noleak", "dl-x-url-png-resnap",
            "dl-x-url-hidden", "dl-x-url-again",
            "dl-x-ref-attr", "dl-x-ref-bytes", "dl-x-ref-noleak", "dl-x-ref-rel", "dl-x-ref-resnap",
            "dl-x-ref-export", "dl-x-ref-saveas", "dl-x-ref-keeps",
            "dl-x-ref-hidden", "dl-x-ref-iframe", "dl-x-ref-iframe-has", "dl-x-ref-channel"
        };

        internal static void RunExtra(
            TestRun run,
            NavSession s,
            FixtureHttp http,
            byte[] txtBytes,
            byte[] pngBytes)
        {
            RunExtraNeg(run, s, http);
            RunExtraD18(run, s, http);
            RunExtraUrl(run, s, http, txtBytes, pngBytes);
            RunExtraRef(run, s, http, txtBytes);
        }

        private static void RunExtraNeg(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), ExtraIds[0]);
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 巩固夹具", ExtraAfter(ExtraIds[0]));
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "下载附件", "普通按钮", "示例图", "另存附件");
            string fileRef = FindDlRef(tree, "下载附件");
            string any = fileRef ?? "e1";
            string txtUrl = http.Url("download", "sample.txt");
            string pngUrl = http.Url("download", "sample.png");

            IntCases.ExpectFail(run, "dl-x-neg-mailto", s.Download(id, url: "mailto:a@b.com", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-data", s.Download(id, url: "data:text/plain,hi", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-ftp", s.Download(id, url: "ftp://127.0.0.1/a.txt", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-about", s.Download(id, url: "about:blank", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-js-void", s.Download(id, url: "javascript:void(0)", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-relative", s.Download(id, url: "sample.txt", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-e0", s.Download(id, refId: "e0", path: "browser_dl/x-neg.txt"), "过期");
            IntCases.ExpectFail(run, "dl-x-neg-e00", s.Download(id, refId: "e00", path: "browser_dl/x-neg.txt"), "过期");
            IntCases.ExpectFail(run, "dl-x-neg-both-png", s.Download(id, refId: any, url: pngUrl, path: "browser_dl/x-neg.txt"), "ref");
            IntCases.ExpectFail(run, "dl-x-neg-wrong-kind-ref", s.Download("w1", refId: any, path: "browser_dl/x-neg.txt"), "渠道");
            IntCases.ExpectFail(run, "dl-x-neg-unknown-ref", s.Download("b999", refId: any, path: "browser_dl/x-neg.txt"), "渠道");
            IntCases.ExpectFail(
                run,
                "dl-x-neg-plain2",
                s.Download(id, refId: FindDlRef(tree, "普通按钮") ?? "e1", path: "browser_dl/x-neg.txt"),
                "下载");
            IntCases.ExpectFail(
                run,
                "dl-x-neg-image2",
                s.Download(id, refId: FindDlRef(tree, "示例图") ?? "e1", path: "browser_dl/x-neg.txt"),
                "url");
            string frame = IntCases.FindRef(tree, "附件框") ?? IntCases.FindIframeRef(tree);
            IntCases.ExpectFail(
                run,
                "dl-x-neg-iframe2",
                s.Download(id, refId: string.IsNullOrEmpty(frame) ? "e1" : frame, path: "browser_dl/x-neg.txt"),
                "展开");
            IntCases.ExpectFail(run, "dl-x-neg-file2", s.Download(id, url: "file:///C:/Windows/win.ini", path: "browser_dl/x-neg.txt"), "http");
            IntCases.ExpectFail(run, "dl-x-neg-neither2", s.Download(id, path: "browser_dl/x-neg.txt"), "ref");
            IntCases.ExpectFail(run, "dl-x-neg-nopath2", s.Download(id, url: txtUrl), "path");

            if (string.IsNullOrEmpty(fileRef))
            {
                run.CaseFail("dl-x-neg-stale2", "空拍无下载附件");
            }
            else
            {
                ToolResult jump = s.Nav(http.Url("public-example", "index.html"), host: IntCases.NavHost, channelId: id);
                if (!jump.Success)
                {
                    run.CaseFail("dl-x-neg-stale2", jump.Error ?? "跳转失败");
                }
                else
                {
                    IntCases.ExpectFail(
                        run,
                        "dl-x-neg-stale2",
                        s.Download(id, refId: fileRef, path: "browser_dl/x-neg.txt"),
                        "过期");
                }
            }
        }

        private static void RunExtraD18(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("d18-risk", "index.html"), "dl-x-neg-pay");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 d18-risk 巩固",
                    "dl-x-neg-delete", "dl-x-neg-submit", "dl-x-neg-pay2", "dl-x-neg-signin", "dl-x-neg-delen");
                return;
            }

            string t = IntCases.WaitTree(s, id, true, "立即支付", "删除", "提交", "付款", "Sign in", "Delete");
            IntCases.ExpectFail(run, "dl-x-neg-pay", s.Download(id, refId: IntCases.FindRef(t, "立即支付") ?? "e1", path: "browser_dl/x-neg.txt"), "支付");
            IntCases.ExpectFail(run, "dl-x-neg-delete", s.Download(id, refId: IntCases.FindRef(t, "删除") ?? "e1", path: "browser_dl/x-neg.txt"), "删除");
            IntCases.ExpectFail(run, "dl-x-neg-submit", s.Download(id, refId: IntCases.FindRef(t, "提交") ?? "e1", path: "browser_dl/x-neg.txt"), "提交");
            IntCases.ExpectFail(run, "dl-x-neg-pay2", s.Download(id, refId: IntCases.FindRef(t, "付款") ?? "e1", path: "browser_dl/x-neg.txt"), "支付");
            IntCases.ExpectFail(run, "dl-x-neg-signin", s.Download(id, refId: IntCases.FindRef(t, "Sign in") ?? "e1", path: "browser_dl/x-neg.txt"), "登录");
            IntCases.ExpectFail(run, "dl-x-neg-delen", s.Download(id, refId: IntCases.FindRef(t, "Delete") ?? "e1", path: "browser_dl/x-neg.txt"), "删除");
        }

        private static void RunExtraUrl(
            TestRun run,
            NavSession s,
            FixtureHttp http,
            byte[] txtBytes,
            byte[] pngBytes)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-x-url-txt");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 巩固 url",
                    "dl-x-url-txt-bytes", "dl-x-url-txt-rel", "dl-x-url-txt-noleak", "dl-x-url-txt-resnap",
                    "dl-x-url-txt-source", "dl-x-url-txt-kind", "dl-x-url-txt-track",
                    "dl-x-url-png", "dl-x-url-png-rel", "dl-x-url-png-noleak", "dl-x-url-png-resnap",
                    "dl-x-url-hidden", "dl-x-url-again");
                return;
            }

            ToolResult txt = s.Download(id, url: http.Url("download", "sample.txt"), path: "browser_dl/x-sample.txt");
            bool txtOk = IsOkFile(txt, txtBytes.Length) && WorkspaceHas(s, Filename(txt), "easywrite-dl-sample");
            IntCases.Ok(run, "dl-x-url-txt", txtOk, Why(txt, "巩固 url txt"));
            IntCases.Ok(run, "dl-x-url-txt-bytes", txt != null && txt.Success && BytesOf(txt) == txtBytes.Length, "巩固 url bytes");
            IntCases.Ok(run, "dl-x-url-txt-rel", txt != null && txt.Success && IsRelative(Filename(txt)), Filename(txt));
            IntCases.Ok(run, "dl-x-url-txt-noleak", txt != null && txt.Success && !HasLeak(txt), "巩固 url 泄漏");
            IntCases.Ok(
                run,
                "dl-x-url-txt-resnap",
                txt != null && txt.Success
                    && IntCases.Flag(txt, "resnapshot")
                    && !IntCases.Flag(txt, "refs_invalidated")
                    && IntCases.Tree(txt).Length > 0,
                "巩固 url 未回树");
            IntCases.Ok(
                run,
                "dl-x-url-txt-source",
                txt != null && txt.Success && IntCases.Str(txt, "source_url").IndexOf("sample.txt", System.StringComparison.Ordinal) >= 0,
                IntCases.Str(txt, "source_url"));
            IntCases.Ok(run, "dl-x-url-txt-kind", txt != null && txt.Success && IntCases.Str(txt, "kind") == "browser", IntCases.Str(txt, "kind"));
            IntCases.Ok(run, "dl-x-url-txt-track", txt != null && txt.Success && IntCases.Str(txt, "track") == "agent", IntCases.Str(txt, "track"));

            ToolResult png = s.Download(id, url: http.Url("download", "sample.png"), path: "browser_dl/x-sample.png");
            IntCases.Ok(
                run,
                "dl-x-url-png",
                IsOkFile(png, pngBytes.Length) || (png != null && png.Success && BytesOf(png) > 0),
                Why(png, "巩固 url png"));
            IntCases.Ok(run, "dl-x-url-png-rel", png != null && png.Success && IsRelative(Filename(png)), Filename(png));
            IntCases.Ok(run, "dl-x-url-png-noleak", png != null && png.Success && !HasLeak(png), "巩固 png 泄漏");
            IntCases.Ok(
                run,
                "dl-x-url-png-resnap",
                png != null && png.Success
                    && IntCases.Flag(png, "resnapshot")
                    && !IntCases.Flag(png, "refs_invalidated")
                    && IntCases.Tree(png).Length > 0,
                "巩固 png 未回树");

            ToolResult again = s.Download(id, url: http.Url("download", "sample.txt"), path: "browser_dl/x-again.txt");
            IntCases.Ok(
                run,
                "dl-x-url-again",
                again != null && again.Success && IsRelative(Filename(again)) && WorkspaceHas(s, Filename(again), "easywrite-dl-sample"),
                Why(again, "巩固再下"));

            string hid = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-x-url-hidden", visible: false);
            if (hid == null)
            {
                return;
            }

            ToolResult hidden = s.Download(hid, url: http.Url("download", "sample.txt"), path: "browser_dl/x-hidden.txt");
            IntCases.Ok(run, "dl-x-url-hidden", hidden != null && hidden.Success, Why(hidden, "巩固隐藏 url"));
        }

        private static void RunExtraRef(TestRun run, NavSession s, FixtureHttp http, byte[] txtBytes)
        {
            string id = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-x-ref-attr");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 download 巩固 ref",
                    "dl-x-ref-bytes", "dl-x-ref-noleak", "dl-x-ref-rel", "dl-x-ref-resnap",
                    "dl-x-ref-export", "dl-x-ref-saveas", "dl-x-ref-keeps",
                    "dl-x-ref-hidden", "dl-x-ref-iframe", "dl-x-ref-iframe-has", "dl-x-ref-channel");
                return;
            }

            string tree = IntCases.WaitTree(s, id, false, "下载附件", "导出", "另存附件");
            string attr = FindDlRef(tree, "下载附件");
            ToolResult got = string.IsNullOrEmpty(attr)
                ? null
                : s.Download(id, refId: attr, path: "browser_dl/x-ref-attr.txt");
            bool ok = IsOkFile(got, txtBytes.Length) && WorkspaceHas(s, Filename(got), "easywrite-dl-sample");
            IntCases.Ok(run, "dl-x-ref-attr", ok, Why(got, "巩固 ref 附件"));
            IntCases.Ok(run, "dl-x-ref-bytes", got != null && got.Success && BytesOf(got) == txtBytes.Length, "巩固 ref bytes");
            IntCases.Ok(run, "dl-x-ref-noleak", got != null && got.Success && !HasLeak(got), "巩固 ref 泄漏");
            IntCases.Ok(run, "dl-x-ref-rel", got != null && got.Success && IsRelative(Filename(got)), Filename(got));
            IntCases.Ok(
                run,
                "dl-x-ref-resnap",
                got != null && got.Success
                    && IntCases.Flag(got, "resnapshot")
                    && !IntCases.Flag(got, "refs_invalidated")
                    && IntCases.Tree(got).Length > 0,
                "巩固 ref 未回树");
            IntCases.Ok(
                run,
                "dl-x-ref-channel",
                got != null && got.Success && NavSession.IsShortId(IntCases.Str(got, "channel_id")),
                IntCases.Str(got, "channel_id"));

            string after = got != null && got.Success ? IntCases.Tree(got) : tree;
            string export = FindDlRef(after, "导出") ?? FindDlRef(IntCases.WaitTree(s, id, false, "导出"), "导出");
            ToolResult exp = string.IsNullOrEmpty(export)
                ? null
                : s.Download(id, refId: export, path: "browser_dl/x-ref-export.txt");
            IntCases.Ok(run, "dl-x-ref-export", exp != null && exp.Success && IsRelative(Filename(exp)), Why(exp, "巩固导出"));

            string neu = exp != null && exp.Success ? IntCases.Tree(exp) : after;
            IntCases.Ok(
                run,
                "dl-x-ref-keeps",
                !string.IsNullOrEmpty(neu)
                    && (neu.IndexOf("下载附件", System.StringComparison.Ordinal) >= 0
                        || neu.IndexOf("导出", System.StringComparison.Ordinal) >= 0
                        || neu.IndexOf("另存附件", System.StringComparison.Ordinal) >= 0),
                "巩固后树不见下载链");

            string save = FindDlRef(neu, "另存附件") ?? FindDlRef(IntCases.WaitTree(s, id, false, "另存附件"), "另存附件");
            ToolResult saveas = string.IsNullOrEmpty(save)
                ? null
                : s.Download(id, refId: save, path: "browser_dl/x-ref-saveas.txt");
            IntCases.Ok(
                run,
                "dl-x-ref-saveas",
                saveas != null && saveas.Success && WorkspaceHas(s, Filename(saveas), "easywrite-dl-sample"),
                Why(saveas, "巩固另存"));

            string hid = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-x-ref-hidden", visible: false);
            if (hid == null)
            {
                run.CaseFail("dl-x-ref-iframe", "依赖隐藏窗开页");
                run.CaseFail("dl-x-ref-iframe-has", "依赖隐藏窗开页");
                return;
            }

            string ht = IntCases.WaitTree(s, hid, false, "下载附件");
            ToolResult href = s.Download(hid, refId: FindDlRef(ht, "下载附件"), path: "browser_dl/x-ref-hidden.txt");
            IntCases.Ok(run, "dl-x-ref-hidden", href != null && href.Success, Why(href, "巩固隐藏 ref"));

            string ifr = IntCases.OpenOk(run, s, http.Url("download", "index.html"), "dl-x-ref-iframe");
            if (ifr == null)
            {
                run.CaseFail("dl-x-ref-iframe-has", "依赖 iframe 开页");
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
                : s.Download(ifr, refId: note, path: "browser_dl/x-ref-iframe.txt");
            IntCases.Ok(run, "dl-x-ref-iframe", innerDl != null && innerDl.Success, innerDl == null ? "巩固展开后无子页备注" : Why(innerDl, "巩固 iframe"));
            IntCases.Ok(
                run,
                "dl-x-ref-iframe-has",
                innerDl != null && innerDl.Success && WorkspaceHas(s, Filename(innerDl), "easywrite-dl-sample"),
                "巩固子页内容不对");
        }

        private static string[] ExtraAfter(string first)
        {
            var rest = new string[ExtraIds.Length - 1];
            int n = 0;
            for (int i = 0; i < ExtraIds.Length; i++)
            {
                if (ExtraIds[i] != first)
                {
                    rest[n++] = ExtraIds[i];
                }
            }

            return rest;
        }
    }
}
