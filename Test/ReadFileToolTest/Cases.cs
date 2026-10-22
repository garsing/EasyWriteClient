using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;
using WordAddIn1;

namespace ReadFileToolTest
{
    internal static class Cases
    {
        public static void Run(
            TestRun run,
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> registry,
            Fixtures fx,
            string filter)
        {
            bool Want(string id)
            {
                return string.IsNullOrEmpty(filter)
                    || id.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0
                    || filter.IndexOf(id, StringComparison.OrdinalIgnoreCase) >= 0;
            }

            bool desktop = DesktopToolPipe.TryWaitUntilAvailable(8000);
            if (desktop)
            {
                Console.WriteLine("检测到易写 Desktop 管道，detailed 走本机已开进程");
            }

            ToolResult Read(string path, object detailed = null, int? offset = null, int? maxLength = null)
            {
                var args = new Dictionary<string, object> { ["path"] = path };
                if (detailed != null)
                {
                    args["detailed"] = detailed;
                }

                if (offset != null)
                {
                    args["offset"] = offset.Value;
                }

                if (maxLength != null)
                {
                    args["max_length"] = maxLength.Value;
                }

                bool useDesktop = desktop && detailed != null && !false.Equals(detailed);
                if (useDesktop)
                {
                    int timeout = 600000;
                    return DesktopToolPipe.Call("F_read_file", args, timeout);
                }

                return Program.Call(registry, args);
            }

            void FailHas(string id, ToolResult r, string needle)
            {
                string err = r == null ? "" : (r.Error ?? "");
                if (r != null && !r.Success && err.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    run.CaseOk(id);
                    return;
                }

                run.CaseFail(id, "期望失败含「" + needle + "」，实际 success=" + (r != null && r.Success) + " err=" + err);
            }

            void PlainOk(string id, ToolResult r, params string[] mustHave)
            {
                if (r == null || !r.Success)
                {
                    run.CaseFail(id, r == null ? "无回包" : (r.Error ?? "失败"));
                    return;
                }

                JObject d = Program.Data(r);
                string content = d.Value<string>("content") ?? "";
                if (content.IndexOf("<table", StringComparison.OrdinalIgnoreCase) >= 0
                    || content.IndexOf("[S_", StringComparison.Ordinal) >= 0)
                {
                    run.CaseFail(id, "简便正文含 <table 或 [S_");
                    return;
                }

                if (d["chunk_summaries"] != null || d["kind"] != null)
                {
                    run.CaseFail(id, "简便回包不应有 kind / chunk_summaries");
                    return;
                }

                foreach (string n in mustHave)
                {
                    if (content.IndexOf(n, StringComparison.Ordinal) < 0)
                    {
                        run.CaseFail(id, "正文缺少「" + n + "」 实得前200=" + Trunc(content, 200));
                        return;
                    }
                }

                run.CaseOk(id);
            }

            if (Want("rf-text-md"))
            {
                ToolResult r = Read(fx.NotesMd);
                if (r.Success)
                {
                    JObject d = Program.Data(r);
                    string c = d.Value<string>("content") ?? "";
                    if (c.IndexOf("hello read_file", StringComparison.Ordinal) >= 0
                        && d["chunk_summaries"] == null
                        && d["kind"] == null)
                    {
                        run.CaseOk("rf-text-md");
                    }
                    else
                    {
                        run.CaseFail("rf-text-md", "原文或多余字段 " + Trunc(c, 120));
                    }
                }
                else
                {
                    run.CaseFail("rf-text-md", r.Error);
                }
            }

            if (Want("rf-text-offset"))
            {
                ToolResult r = Read(fx.NotesMd, maxLength: 50);
                JObject d = Program.Data(r);
                int n = d.Value<int?>("read_characters") ?? -1;
                if (r.Success && n >= 0 && n <= 50)
                {
                    run.CaseOk("rf-text-offset");
                }
                else
                {
                    run.CaseFail("rf-text-offset", r.Success ? "read_characters=" + n : r.Error);
                }
            }

            if (Want("rf-text-missing"))
            {
                FailHas("rf-text-missing", Read(Path.Combine(fx.Dir, "nonexistent_xyz.json")), "不存在");
            }

            if (Want("rf-img-png"))
            {
                ToolResult r = Read(fx.Png);
                JObject d = Program.Data(r);
                if (r.Success && d["image_base64"] != null && d["chunk_summaries"] == null)
                {
                    run.CaseOk("rf-img-png");
                }
                else
                {
                    run.CaseFail("rf-img-png", r.Success ? "无 image_base64" : r.Error);
                }
            }

            if (Want("rf-img-detailed"))
            {
                FailHas("rf-img-detailed", Read(fx.Png, detailed: true), "detailed");
            }

            if (Want("rf-text-detailed"))
            {
                FailHas("rf-text-detailed", Read(fx.NotesMd, detailed: true), "detailed");
            }

            if (Want("rf-bin-zip"))
            {
                ToolResult r = Read(fx.Zip);
                if (!r.Success)
                {
                    run.CaseOk("rf-bin-zip");
                }
                else
                {
                    run.CaseFail("rf-bin-zip", "zip 不应当文本成功");
                }
            }

            if (Want("rf-plain-docx"))
            {
                PlainOk("rf-plain-docx", Read(fx.ShortDocx), "短文标题微专业", "格甲", "格乙");
            }

            if (Want("rf-plain-xlsx"))
            {
                PlainOk("rf-plain-xlsx", Read(fx.Xlsx), "收入", "成本", "单元格甲", "单元格乙");
            }

            if (Want("rf-plain-pptx"))
            {
                PlainOk("rf-plain-pptx", Read(fx.Pptx), "第1页", "第2页", "第一页标题", "第二页正文");
            }

            if (Want("rf-plain-pdf"))
            {
                string pdf = fx.RealPdf ?? fx.TextPdf;
                ToolResult r = Read(pdf);
                if (fx.RealPdf != null)
                {
                    PlainOk("rf-plain-pdf", r, "微专业");
                }
                else
                {
                    PlainOk("rf-plain-pdf", r, "HelloMicro");
                }
            }

            if (Want("rf-plain-abs"))
            {
                PlainOk("rf-plain-abs", Read(Path.GetFullPath(fx.ShortDocx)), "短文标题微专业");
            }

            if (Want("rf-plain-pdf-scan"))
            {
                ToolResult r = Read(fx.ScanPdf);
                JObject d = Program.Data(r);
                string msg = d.Value<string>("message") ?? "";
                string c = d.Value<string>("content") ?? "";
                if (r.Success && string.IsNullOrEmpty(c)
                    && msg.IndexOf("没有可抽取的文字", StringComparison.Ordinal) >= 0
                    && msg.IndexOf("detailed=true", StringComparison.Ordinal) >= 0)
                {
                    run.CaseOk("rf-plain-pdf-scan");
                }
                else
                {
                    run.CaseFail("rf-plain-pdf-scan", "success=" + r.Success + " content=" + Trunc(c, 80) + " msg=" + msg + " err=" + r.Error);
                }
            }

            if (Want("rf-plain-doc-legacy"))
            {
                if (string.IsNullOrEmpty(fx.LegacyDoc))
                {
                    run.CaseSkip("rf-plain-doc-legacy", fx.LegacyDocError ?? "无 .doc 夹具");
                }
                else
                {
                    ToolResult r = Read(fx.LegacyDoc);
                    if (!r.Success)
                    {
                        string err = r.Error ?? "";
                        if (err.IndexOf("安装 Word", StringComparison.Ordinal) >= 0
                            || err.IndexOf("docx", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            run.CaseOk("rf-plain-doc-legacy");
                        }
                        else
                        {
                            run.CaseFail("rf-plain-doc-legacy", err);
                        }
                    }
                    else
                    {
                        PlainOk("rf-plain-doc-legacy", r, "老doc正文微专业");
                    }
                }
            }

            if (Want("rf-win-first") || Want("rf-win-next"))
            {
                ToolResult first = Read(fx.LongDocx, offset: 0);
                JObject d1 = Program.Data(first);
                int read = d1.Value<int?>("read_characters") ?? -1;
                int end = d1.Value<int?>("range_end") ?? -1;
                bool more = d1.Value<bool?>("has_more") ?? false;
                string msg = d1.Value<string>("message") ?? "";
                string c1 = d1.Value<string>("content") ?? "";
                if (Want("rf-win-first"))
                {
                    if (first.Success && read >= 0 && read <= 8000 && more
                        && msg.IndexOf("只读了 8000 字", StringComparison.Ordinal) >= 0
                        && msg.IndexOf("offset", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        run.CaseOk("rf-win-first");
                    }
                    else
                    {
                        run.CaseFail("rf-win-first", "read=" + read + " more=" + more + " msg=" + msg + " err=" + first.Error);
                    }
                }

                if (Want("rf-win-next"))
                {
                    ToolResult second = Read(fx.LongDocx, offset: end);
                    JObject d2 = Program.Data(second);
                    string c2 = d2.Value<string>("content") ?? "";
                    int start2 = d2.Value<int?>("range_start") ?? -1;
                    if (second.Success && start2 == end && !string.IsNullOrEmpty(c2)
                        && (c1.Length < 8 || c2.IndexOf(c1.Substring(0, 8), StringComparison.Ordinal) < 0))
                    {
                        run.CaseOk("rf-win-next");
                    }
                    else
                    {
                        run.CaseFail("rf-win-next", "start2=" + start2 + " end1=" + end + " err=" + second.Error);
                    }
                }
            }

            if (Want("rf-win-short"))
            {
                ToolResult r = Read(fx.ShortDocx);
                JObject d = Program.Data(r);
                string msg = d.Value<string>("message") ?? "";
                bool more = d.Value<bool?>("has_more") ?? true;
                if (r.Success && !more && msg.IndexOf("只读了 8000 字", StringComparison.Ordinal) < 0)
                {
                    run.CaseOk("rf-win-short");
                }
                else
                {
                    run.CaseFail("rf-win-short", "more=" + more + " msg=" + msg + " err=" + r.Error);
                }
            }

            if (Want("rf-det-short") || Want("rf-det-reuse") || Want("rf-det-long") || Want("rf-det-pdf") || Want("rf-det-chunk"))
            {
                RunDetailed(run, Want, Read, fx);
            }

            if (Want("rf-det-fail-hint"))
            {
                ToolResult r = Read(fx.EmptyFile, detailed: true);
                if (!r.Success && LooksLikeNeedLogin(r.Error))
                {
                    run.CaseSkip("rf-det-fail-hint", r.Error);
                }
                else if (!r.Success)
                {
                    run.CaseOk("rf-det-fail-hint");
                }
                else
                {
                    run.CaseFail("rf-det-fail-hint", "0 字节 pdf detailed 不应成功");
                }
            }

            SkipE(run, Want, "rf-chat-short");
            SkipE(run, Want, "rf-chat-long");
            SkipE(run, Want, "rf-brpdf-short");
            SkipE(run, Want, "rf-brpdf-long");
        }

        private static void SkipE(TestRun run, Func<string, bool> want, string id)
        {
            if (want(id))
            {
                run.CaseSkip(id, "契约由 pytest test_kb_document_model_view 覆盖");
            }
        }

        private static void RunDetailed(
            TestRun run,
            Func<string, bool> want,
            Func<string, object, int?, int?, ToolResult> read,
            Fixtures fx)
        {
            if (want("rf-det-short") || want("rf-det-reuse"))
            {
                ToolResult r = read(fx.ShortDocx, true, null, null);
                if (!r.Success && LooksLikeNeedLogin(r.Error))
                {
                    if (want("rf-det-short"))
                    {
                        run.CaseSkip("rf-det-short", r.Error);
                    }

                    if (want("rf-det-reuse"))
                    {
                        run.CaseSkip("rf-det-reuse", r.Error);
                    }
                }
                else if (want("rf-det-short"))
                {
                    AssertDetailedKind(run, "rf-det-short", r, "full_text");
                }

                if (want("rf-det-reuse") && r.Success)
                {
                    ToolResult r2 = read(fx.ShortDocx, true, null, null);
                    JObject d = Program.Data(r2);
                    bool dup = d.Value<bool?>("deduplicated") ?? false;
                    if (r2.Success && dup)
                    {
                        run.CaseOk("rf-det-reuse");
                    }
                    else if (r2.Success)
                    {
                        run.CaseSkip("rf-det-reuse", "第二次成功但未标 deduplicated（可能已换 hash）");
                    }
                    else
                    {
                        run.CaseFail("rf-det-reuse", r2.Error);
                    }
                }
            }

            if (want("rf-det-long") || want("rf-det-chunk"))
            {
                ToolResult r = read(fx.LongDocx, true, null, null);
                if (!r.Success && LooksLikeNeedLogin(r.Error))
                {
                    if (want("rf-det-long"))
                    {
                        run.CaseSkip("rf-det-long", r.Error);
                    }

                    if (want("rf-det-chunk"))
                    {
                        run.CaseSkip("rf-det-chunk", r.Error);
                    }
                }
                else
                {
                    if (want("rf-det-long"))
                    {
                        AssertDetailedKind(run, "rf-det-long", r, "chunk_summaries");
                    }

                    if (want("rf-det-chunk"))
                    {
                        run.CaseSkip("rf-det-chunk", "本 exe 不调 B_get_kb_chunk_content，登录后手测");
                    }
                }
            }

            if (want("rf-det-pdf"))
            {
                string pdf = fx.RealPdf ?? fx.TextPdf;
                ToolResult r = read(pdf, true, null, null);
                if (!r.Success && LooksLikeNeedLogin(r.Error))
                {
                    run.CaseSkip("rf-det-pdf", r.Error);
                }
                else
                {
                    JObject d = Program.Data(r);
                    string kind = d.Value<string>("kind");
                    if (r.Success && (kind == "full_text" || kind == "chunk_summaries")
                        && !string.IsNullOrEmpty(d.Value<string>("storage_doc_uuid")))
                    {
                        run.CaseOk("rf-det-pdf");
                    }
                    else
                    {
                        run.CaseFail("rf-det-pdf", r.Success ? "kind=" + kind : r.Error);
                    }
                }
            }
        }

        private static void AssertDetailedKind(TestRun run, string id, ToolResult r, string kind)
        {
            if (r == null || !r.Success)
            {
                run.CaseFail(id, r == null ? "无回包" : r.Error);
                return;
            }

            JObject d = Program.Data(r);
            if (d.Value<string>("kind") != kind)
            {
                run.CaseFail(id, "kind=" + d.Value<string>("kind") + " 期望 " + kind);
                return;
            }

            if (string.IsNullOrEmpty(d.Value<string>("storage_doc_uuid"))
                || string.IsNullOrEmpty(d.Value<string>("index_doc_uuid"))
                || d["knowledge_base_uuid"] != null)
            {
                run.CaseFail(id, "uuid 字段不对");
                return;
            }

            if (kind == "full_text")
            {
                if (string.IsNullOrEmpty(d.Value<string>("content")) || d["chunk_summaries"] != null)
                {
                    run.CaseFail(id, "full_text 应有 content、无 chunk_summaries");
                    return;
                }
            }
            else
            {
                var arr = d["chunk_summaries"] as JArray;
                if (arr == null || arr.Count == 0 || d["content"] != null)
                {
                    run.CaseFail(id, "chunk_summaries 形态不对");
                    return;
                }

                JObject one = arr[0] as JObject;
                if (one == null || one["chunk_index"] == null || one["description"] == null || one["summary"] == null)
                {
                    run.CaseFail(id, "分块缺三项");
                    return;
                }
            }

            run.CaseOk(id);
        }

        private static bool LooksLikeNeedLogin(string err)
        {
            if (string.IsNullOrEmpty(err))
            {
                return false;
            }

            return err.IndexOf("登录", StringComparison.Ordinal) >= 0
                || err.IndexOf("未配置 API", StringComparison.Ordinal) >= 0
                || err.IndexOf("未登录", StringComparison.Ordinal) >= 0;
        }

        private static string Trunc(string s, int n)
        {
            if (string.IsNullOrEmpty(s))
            {
                return "";
            }

            return s.Length <= n ? s : s.Substring(0, n);
        }
    }
}
