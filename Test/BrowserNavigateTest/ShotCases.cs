using System;
using System.Collections.Generic;
using WordAddIn1;

namespace BrowserNavigateTest
{
    /// <summary>第 6 波网页可视区截图，只测 agent。</summary>
    internal static class ShotCases
    {
        public static void Run(TestRun run, FixtureHttp http, int chromeBefore, int edgeBefore)
        {
            var s = new NavSession();
            RunNeg(run, s, http);
            RunOk(run, s, http);
            RunScroll(run, s, http);
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
            string id = IntCases.OpenOk(run, s, http.Url("viewport-shot", "index.html"), "shot-neg-wrong-kind");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 viewport-shot 夹具",
                    "shot-neg-unknown", "shot-neg-page-number", "shot-neg-sheet", "shot-neg-range",
                    "shot-neg-sheet-range", "shot-neg-page-from", "shot-neg-page-to", "shot-neg-from-to",
                    "shot-neg-shape-id", "shot-neg-shape-ids", "shot-neg-slide-id", "shot-neg-page-numbers",
                    "shot-neg-storage", "shot-neg-source-channel", "shot-neg-doc-name", "shot-neg-path-escape");
                return;
            }

            IntCases.ExpectFail(run, "shot-neg-wrong-kind", s.Capture("w1"), "channel_id");
            IntCases.ExpectFail(run, "shot-neg-unknown", s.Capture("b999"), "channel_id");
            IntCases.ExpectFail(run, "shot-neg-page-number", s.Capture(id, pageNumber: 1), "page_number");
            IntCases.ExpectFail(run, "shot-neg-sheet", s.Capture(id, sheet: "Sheet1"), "sheet");
            IntCases.ExpectFail(run, "shot-neg-range", s.Capture(id, range: "A1:B2"), "range");
            IntCases.ExpectFail(run, "shot-neg-sheet-range", s.Capture(id, sheet: "Sheet1", range: "A1:B2"), "sheet");
            IntCases.ExpectFail(run, "shot-neg-page-from", s.Capture(id, pageFrom: 1), "page_from");
            IntCases.ExpectFail(run, "shot-neg-page-to", s.Capture(id, pageTo: 2), "page_to");
            IntCases.ExpectFail(run, "shot-neg-from-to", s.Capture(id, pageFrom: 1, pageTo: 2), "page_from");
            IntCases.ExpectFail(run, "shot-neg-shape-id", s.Capture(id, shapeId: "s1"), "F_capture_ppt_shape");
            IntCases.ExpectFail(run, "shot-neg-shape-ids", s.Capture(id, shapeIds: "s1,s2"), "F_capture_ppt_shape");
            IntCases.ExpectFail(run, "shot-neg-slide-id", s.Capture(id, slideId: "256"), "slide_id");
            IntCases.ExpectFail(run, "shot-neg-page-numbers", s.Capture(id, pageNumbers: "1,2"), "page_numbers");
            IntCases.ExpectFail(run, "shot-neg-storage", s.Capture(id, storageDocUuid: "kb-uuid"), "storage_doc_uuid");
            IntCases.ExpectFail(run, "shot-neg-source-channel", s.Capture(id, sourceChannelId: id), "source_channel_id");
            IntCases.ExpectFail(run, "shot-neg-doc-name", s.Capture(id, documentName: "foo.docx"), "名");
            IntCases.ExpectFail(
                run,
                "shot-neg-path-escape",
                s.Capture(id, path: "../outside.jpg"),
                "会话");
        }

        private static void RunOk(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("viewport-shot", "index.html"), "shot-ok-has-image");
            if (id == null)
            {
                IntCases.FailMany(run, "依赖 viewport-shot 夹具",
                    "shot-ok-kind", "shot-ok-channel", "shot-ok-url", "shot-ok-title",
                    "shot-ok-size", "shot-ok-format", "shot-ok-twice", "shot-ok-hidden",
                    "shot-ok-path", "shot-ok-path-has-b64", "shot-snap-no-image");
                return;
            }

            ToolResult first = s.Capture(id);
            IntCases.Ok(run, "shot-ok-has-image", HasImage(first), Why(first, "可视区无图"));
            IntCases.Ok(run, "shot-ok-kind", first != null && first.Success && IntCases.Str(first, "source_kind") == "browser", IntCases.Str(first, "source_kind"));
            IntCases.Ok(run, "shot-ok-channel", first != null && first.Success && NavSession.IsShortId(IntCases.Str(first, "channel_id")), IntCases.Str(first, "channel_id"));
            IntCases.Ok(
                run,
                "shot-ok-url",
                first != null && first.Success && IntCases.Str(first, "url").IndexOf("viewport-shot", StringComparison.OrdinalIgnoreCase) >= 0,
                IntCases.Str(first, "url"));
            IntCases.Ok(
                run,
                "shot-ok-title",
                first != null && first.Success && IntCases.Str(first, "title").IndexOf("可视区", StringComparison.Ordinal) >= 0,
                IntCases.Str(first, "title"));
            IntCases.Ok(run, "shot-ok-size", Width(first) > 8 && Height(first) > 8, Width(first) + "x" + Height(first));
            IntCases.Ok(
                run,
                "shot-ok-format",
                first != null && first.Success && IsJpegOrPng(IntCases.Str(first, "format")),
                IntCases.Str(first, "format"));

            ToolResult twice = s.Capture(id);
            IntCases.Ok(run, "shot-ok-twice", HasImage(twice), Why(twice, "再截"));

            ToolResult snap = s.Snap(id);
            IntCases.Ok(
                run,
                "shot-snap-no-image",
                snap != null && snap.Success && string.IsNullOrEmpty(IntCases.Str(snap, "image_base64")),
                "snapshot 不该带图");

            ToolResult withPath = s.Capture(id, path: "browser_shot/view.jpg");
            string saved = IntCases.Str(withPath, "path");
            IntCases.Ok(
                run,
                "shot-ok-path",
                withPath != null && withPath.Success && IsRelative(saved) && saved.IndexOf("browser_shot", StringComparison.OrdinalIgnoreCase) >= 0,
                saved);
            IntCases.Ok(run, "shot-ok-path-has-b64", HasImage(withPath), "落盘后少了 image_base64");

            string hid = IntCases.OpenOk(run, s, http.Url("viewport-shot", "index.html"), "shot-ok-hidden", visible: false);
            if (hid == null)
            {
                return;
            }

            ToolResult hidden = s.Capture(hid);
            IntCases.Ok(run, "shot-ok-hidden", HasImage(hidden), Why(hidden, "隐藏窗"));
        }

        private static void RunScroll(TestRun run, NavSession s, FixtureHttp http)
        {
            string id = IntCases.OpenOk(run, s, http.Url("viewport-shot", "index.html"), "shot-scroll-changes");
            if (id == null)
            {
                return;
            }

            ToolResult top = s.Capture(id);
            if (!HasImage(top))
            {
                run.CaseFail("shot-scroll-changes", Why(top, "滚动前无图"));
                return;
            }

            for (int i = 0; i < 8; i++)
            {
                ToolResult sc = s.Interact(id, "scroll", direction: "down");
                if (sc == null || !sc.Success)
                {
                    run.CaseFail("shot-scroll-changes", sc == null ? "scroll 无回包" : sc.Error);
                    return;
                }
            }

            ToolResult bottom = s.Capture(id);
            IntCases.Ok(
                run,
                "shot-scroll-changes",
                HasImage(bottom) && !SameImage(top, bottom),
                bottom == null || !bottom.Success ? Why(bottom, "滚动后") : "滚动后图未变（仍像只截到顶区）");
        }

        private static bool HasImage(ToolResult r)
        {
            if (r == null || !r.Success)
            {
                return false;
            }

            string b64 = IntCases.Str(r, "image_base64");
            if (string.IsNullOrWhiteSpace(b64) || b64.Length < 80)
            {
                return false;
            }

            try
            {
                byte[] raw = Convert.FromBase64String(b64);
                return raw != null && raw.Length > 32;
            }
            catch
            {
                return false;
            }
        }

        private static bool SameImage(ToolResult a, ToolResult b)
        {
            if (a == null || b == null)
            {
                return false;
            }

            return string.Equals(IntCases.Str(a, "image_base64"), IntCases.Str(b, "image_base64"), StringComparison.Ordinal);
        }

        private static int Width(ToolResult r)
        {
            return IntOf(r, "width_px");
        }

        private static int Height(ToolResult r)
        {
            return IntOf(r, "height_px");
        }

        private static int IntOf(ToolResult r, string key)
        {
            object v = null;
            Dictionary<string, object> d = IntCases.Data(r);
            if (d == null || !d.TryGetValue(key, out v) || v == null)
            {
                return 0;
            }

            try
            {
                return Convert.ToInt32(v);
            }
            catch
            {
                return 0;
            }
        }

        private static bool IsJpegOrPng(string format)
        {
            string f = (format ?? "").Trim().ToLowerInvariant();
            return f == "jpeg" || f == "jpg" || f == "png";
        }

        private static bool IsRelative(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }

            return path.IndexOf(":\\", StringComparison.Ordinal) < 0
                && !path.StartsWith("\\\\", StringComparison.Ordinal);
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

            return what + " " + Width(r) + "x" + Height(r) + " " + IntCases.Str(r, "format");
        }
    }
}
