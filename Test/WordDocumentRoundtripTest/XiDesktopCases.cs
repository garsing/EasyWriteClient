using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>经 Desktop 测 F_render_svg + F_excel_image insert。清单见 docs/测试/excel测试/贴图。</summary>
    internal static class XiDesktopCases
    {
        public const int Target = 13;
        private static readonly Regex ImageIdRx = new Regex(@"^I_[A-Za-z0-9]{8}$", RegexOptions.Compiled);
        private const string MiniSvg =
            "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"><rect width=\"10\" height=\"10\" fill=\"#09c\"/></svg>";

        public static void Run(TestRun run, string host)
        {
            var s = new OfficeCkptSession(OfficeCkptSession.NormalizeExcelHost(host));
            CkptCaseUtil.ExpectFail(
                run,
                "xid-neg-svg-none",
                Render(null, null),
                "二选一");
            CkptCaseUtil.ExpectFail(
                run,
                "xid-neg-svg-both",
                Render(MiniSvg, "a.svg"),
                "二选一");

            ToolResult rendered = Render(MiniSvg, null);
            string filename = WdSession.Str(rendered, "filename");
            if (string.IsNullOrEmpty(filename))
            {
                filename = WdSession.Str(rendered, "path");
            }

            bool png = !string.IsNullOrEmpty(filename)
                && filename.EndsWith(".png", StringComparison.OrdinalIgnoreCase);
            CkptCaseUtil.Ok(
                run,
                "xid-render-ok",
                rendered.Success && png,
                rendered.Success ? "filename=" + filename : (rendered.Error ?? ""));

            string[] insertIds = { "xid-render-insert", "xid-list-after" };
            string path, id;
            if (CkptCaseUtil.TryOpen(run, s, out path, out id, insertIds))
            {
                ToolResult added = s.AddSheet(id, "贴图");
                if (!added.Success)
                {
                    CkptCaseUtil.FailMany(run, added.Error ?? "加表失败", insertIds);
                    s.CloseByPath(path, id);
                }
                else if (!rendered.Success || !png)
                {
                    CkptCaseUtil.FailMany(run, "无 PNG 不能 insert", insertIds);
                    s.CloseByPath(path, id);
                }
                else
                {
                    ToolResult ins = Image(id, "insert", "贴图", "B2", filename, null, null);
                    string imageId = WdSession.Str(ins, "image_id");
                    CkptCaseUtil.Ok(
                        run,
                        "xid-render-insert",
                        ins.Success && ImageIdRx.IsMatch(imageId),
                        ins.Success ? "image_id=" + imageId : (ins.Error ?? ""));

                    ToolResult listed = Image(id, "list", "贴图", null, null, null, null);
                    int count = WdSession.IntOf(listed, "images_count");
                    CkptCaseUtil.Ok(
                        run,
                        "xid-list-after",
                        listed.Success && count >= 1 && CkptCaseUtil.BlobHas(listed, imageId),
                        listed.Success ? "count=" + count : (listed.Error ?? ""));
                    s.CloseByPath(path, id);
                }
            }

            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-neg-script",
                Render("<svg xmlns=\"http://www.w3.org/2000/svg\"><script>alert(1)</script></svg>", null, null),
                "script");
            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-neg-empty",
                Render("   ", null, null),
                "二选一");
            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-neg-slash",
                Render(MiniSvg, null, "a/b.png"),
                "裸名");
            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-neg-ext",
                Render(MiniSvg, null, "a.jpg"),
                ".png");
            string pngPath = Path.Combine(Path.GetTempPath(), "ew-xid-" + Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(
                pngPath,
                Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="));
            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-neg-png-path",
                Render(null, pngPath, null),
                "svg");

            string named = "xid-" + Guid.NewGuid().ToString("N").Substring(0, 8) + ".png";
            ToolResult namedOk = Render(MiniSvg, null, named);
            CkptCaseUtil.Ok(
                run,
                "xid-g-filename",
                namedOk.Success && string.Equals(WdSession.Str(namedOk, "filename"), named, StringComparison.OrdinalIgnoreCase),
                namedOk.Success ? "filename=" + WdSession.Str(namedOk, "filename") : (namedOk.Error ?? ""));
            CkptCaseUtil.ExpectFail(
                run,
                "xid-g-filename-exists",
                Render(MiniSvg, null, named),
                "已存在");

            string[] namedInsert = { "xid-g-insert-named" };
            string path2, id2;
            if (!CkptCaseUtil.TryOpen(run, s, out path2, out id2, namedInsert))
            {
                return;
            }

            ToolResult add2 = s.AddSheet(id2, "贴图");
            if (!add2.Success || !namedOk.Success)
            {
                CkptCaseUtil.FailMany(run, add2.Error ?? namedOk.Error ?? "铺失败", namedInsert);
                s.CloseByPath(path2, id2);
                return;
            }

            ToolResult ins2 = Image(id2, "insert", "贴图", "B2", named, null, null);
            CkptCaseUtil.Ok(
                run,
                "xid-g-insert-named",
                ins2.Success && ImageIdRx.IsMatch(WdSession.Str(ins2, "image_id")),
                ins2.Success ? "image_id=" + WdSession.Str(ins2, "image_id") : (ins2.Error ?? ""));
            s.CloseByPath(path2, id2);
        }

        private static ToolResult Render(string svg, string path)
        {
            return Render(svg, path, null);
        }

        private static ToolResult Render(string svg, string path, string filename)
        {
            var args = new Dictionary<string, object>();
            if (svg != null)
            {
                args["svg"] = svg;
            }

            if (path != null)
            {
                args["path"] = path;
            }

            if (filename != null)
            {
                args["filename"] = filename;
            }

            return DesktopToolPipe.Call("F_render_svg", args, 180000);
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

            return DesktopToolPipe.Call("F_excel_image", args);
        }
    }
}
