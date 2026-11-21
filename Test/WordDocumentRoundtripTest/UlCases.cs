using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>统一定位 ul-*。对照 docs/测试/word测试/统一定位。</summary>
    internal static class UlCases
    {
        public const int Target = 32;

        private static readonly string[] HeadingNames = { "标题 1", "heading 1", "Heading 1" };
        private static readonly Regex SentMark = new Regex(@"\[(S_[0-9A-Za-z]{5})-start(?: 歧义无法操作)?\]", RegexOptions.Compiled);
        private static readonly Regex ParaMark = new Regex(@"\[(P_[0-9A-Za-z_]+)(?:-start| start)", RegexOptions.Compiled);
        private static readonly Regex CellMark = new Regex(@"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\]", RegexOptions.Compiled);
        private static readonly Regex TableMark = new Regex(@"<table id=""?(T_[0-9A-Za-z]+)""?>", RegexOptions.Compiled);
        private static readonly Regex ImageMark = new Regex(@"<image[^>]*id=""?(I_[0-9A-Za-z]+)""?", RegexOptions.Compiled);

        private static HashSet<string> Only;
        private static readonly HashSet<string> Seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public static void Run(TestRun run, string host, string filter)
        {
            Seen.Clear();
            Only = null;
            if (!string.IsNullOrWhiteSpace(filter)
                && !filter.Equals("ul", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("ul-", StringComparison.OrdinalIgnoreCase))
            {
                var picked = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string part in filter.Split(','))
                {
                    string id = part.Trim();
                    if (id.Length > 0)
                    {
                        picked.Add(id);
                    }
                }

                if (picked.Count > 0)
                {
                    Only = picked;
                    run.TotalTarget = picked.Count;
                }
            }

            var s = new WdSession(host);
            RunMiss(run, s);
            RunToc(run, s);
            RunGeo(run, s);
            RunTable(run, s);
            RunCallers(run, s);
            RunOut(run, s);
            foreach (string id in AllIds())
            {
                if (!Want(id) || Seen.Contains(id))
                {
                    continue;
                }

                Fail(run, id, "测试方法没有执行这条");
            }
        }

        private static string[] AllIds()
        {
            return new[]
            {
                "ul-miss-rest", "ul-order", "ul-parse-rest", "ul-para-miss",
                "ul-toc-block",
                "ul-cell-write", "ul-cell-miss", "ul-cell-mix", "ul-cell-and-sent",
                "ul-table-miss", "ul-table-toc", "ul-table-gone", "ul-del-image", "ul-del-table", "ul-chart-gone",
                "ul-replace-span", "ul-insert-end", "ul-insert-table", "ul-comment-sent", "ul-comment-table",
                "ul-apply-part", "ul-apply-toc", "ul-toc-para", "ul-toc-start", "ul-toc-inside",
                "ul-text-slice", "ul-text-two", "ul-para-span", "ul-apply-ok",
                "ul-blank", "ul-del-comment", "ul-comment-cursor"
            };
        }

        private static void RunMiss(TestRun run, WdSession s)
        {
            if (Want("ul-miss-rest"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-miss-rest", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string yi = Sent(body, "乙句定位");
                        ToolResult wr = s.ProcessActions(id, Batch(
                            Replace("S_xxxxx", "不应出现。"),
                            Replace(yi, "乙句已改。")));
                        string after = Body(s, id);
                        Ok(run, "ul-miss-rest",
                            !ItemOk(wr, 0) && ItemOk(wr, 1)
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("乙句已改。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-order"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。\r丙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-order", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string jia = Sent(body, "甲句定位");
                        string bing = Sent(body, "丙句定位");
                        ToolResult wr = s.ProcessActions(id, Batch(
                            Replace(jia, "甲改。"),
                            Replace("S_xxxxx", "不应出现。"),
                            Replace(bing, "丙改。")));
                        string after = Body(s, id);
                        Ok(run, "ul-order",
                            ItemOk(wr, 0) && !ItemOk(wr, 1) && ItemOk(wr, 2)
                            && after.IndexOf("甲改。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("乙句定位。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("丙改。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-parse-rest"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-parse-rest", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string jia = Sent(body, "甲句定位");
                        string yi = Sent(body, "乙句定位");
                        ToolResult wr = s.ProcessActions(id, Batch(
                            Replace(jia + ",T_xxxxx", "不应出现。"),
                            Replace(yi, "乙句已改。")));
                        string after = Body(s, id);
                        Ok(run, "ul-parse-rest",
                            !ItemOk(wr, 0) && ItemOk(wr, 1)
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("乙句已改。", StringComparison.Ordinal) >= 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-para-miss"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-para-miss", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string yi = Sent(body, "乙句定位");
                        var miss = One("replace", new Dictionary<string, object>
                        {
                            ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = "P_xxxxx" },
                            ["new"] = "不应出现。",
                            ["format"] = DefaultFormat()
                        });
                        ToolResult wr = s.ProcessActions(id, Batch(miss, Replace(yi, "乙句已改。")));
                        string after = Body(s, id);
                        Ok(run, "ul-para-miss",
                            !ItemOk(wr, 0) && ItemOk(wr, 1)
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("乙句已改。", StringComparison.Ordinal) >= 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }
        }

        private static void RunToc(TestRun run, WdSession s)
        {
            if (!Want("ul-toc-block", "ul-table-toc", "ul-apply-toc", "ul-toc-para", "ul-toc-start", "ul-toc-inside"))
            {
                return;
            }

            if (Want("ul-toc-start"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out _))
                {
                    Fail(run, "ul-toc-start", "铺稿失败");
                }
                else
                {
                    try
                    {
                        StyleHeading(s, path);
                        ToolResult toc = Toc(id, "insert", null, "start");
                        string after = Body(s, id);
                        Ok(run, "ul-toc-start",
                            toc.Success
                            && after.IndexOf("[目录域 ", StringComparison.Ordinal) >= 0
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0,
                            Dump(toc) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-toc-para"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-toc-para", "铺稿失败");
                }
                else
                {
                    try
                    {
                        StyleHeading(s, path);
                        string para = Para(ParagraphBody(s, id), "乙句定位");
                        ToolResult toc = Toc(id, "insert", null, para);
                        string after = Body(s, id);
                        int tocAt = after.IndexOf("[目录域 ", StringComparison.Ordinal);
                        int jiaAt = after.IndexOf("甲句定位。", StringComparison.Ordinal);
                        int yiAt = after.IndexOf("乙句定位。", StringComparison.Ordinal);
                        Ok(run, "ul-toc-para",
                            toc.Success && tocAt >= 0 && jiaAt >= 0 && yiAt > tocAt && tocAt < yiAt,
                            Dump(toc) + " para=" + para + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-toc-block", "ul-apply-toc", "ul-toc-inside", "ul-table-toc"))
            {
                if (!Open(s, "标题甲\r域外正文。", out string path, out string id, out _))
                {
                    FailMany(run, "铺目录稿失败", "ul-toc-block", "ul-apply-toc", "ul-toc-inside", "ul-table-toc");
                }
                else
                {
                    try
                    {
                        StyleHeading(s, path);
                        ToolResult inserted = Toc(id, "insert", null, "start");
                        ToolResult got = s.GetContent(id);
                        string interiorP = WdSession.Str(got, "toc_interior_paragraph");
                        string interiorS = WdSession.Str(got, "toc_interior_sentence");
                        string body = got.Success ? WdSession.Str(got, "document_content") : "";
                        string outside = Sent(body, "域外正文");

                        if (Want("ul-toc-block"))
                        {
                            string code = string.IsNullOrEmpty(interiorP) ? interiorS : interiorP;
                            List<Dictionary<string, object>> first;
                            if (!string.IsNullOrEmpty(interiorP))
                            {
                                first = One("replace", new Dictionary<string, object>
                                {
                                    ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = interiorP },
                                    ["new"] = "不应出现。",
                                    ["format"] = DefaultFormat()
                                });
                            }
                            else
                            {
                                first = Replace(string.IsNullOrEmpty(code) ? "S_xxxxx" : code, "不应出现。");
                            }

                            ToolResult wr = s.ProcessActions(id, Batch(first, Replace(outside, "域外已改。")));
                            string after = Body(s, id);
                            Ok(run, "ul-toc-block",
                                !string.IsNullOrEmpty(code)
                                && !wr.Success && WdSession.IntOf(wr, "executed_actions") == 0
                                && after.IndexOf("域外正文。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("域外已改。", StringComparison.Ordinal) < 0
                                && Blob(wr, "目录域里面的编码"),
                                Dump(wr) + " code=" + code + " body=" + Strip(after));
                        }

                        if (Want("ul-apply-toc"))
                        {
                            string code = string.IsNullOrEmpty(interiorP) ? interiorS : interiorP;
                            ToolResult applied = Apply(s, id, code ?? "S_xxxxx", "heading_1");
                            bool styled = StyleHas(s, path, "域外正文。", HeadingNames);
                            Ok(run, "ul-apply-toc",
                                !string.IsNullOrEmpty(code)
                                && !applied.Success
                                && Blob(applied, "目录域里面的编码")
                                && !styled,
                                Dump(applied) + " styled=" + styled);
                        }

                        if (Want("ul-toc-inside"))
                        {
                            int before = Count(Body(s, id), "[目录域 ");
                            ToolResult toc = Toc(id, "insert", null, string.IsNullOrEmpty(interiorP) ? "P_xxxxx" : interiorP);
                            int after = Count(Body(s, id), "[目录域 ");
                            Ok(run, "ul-toc-inside",
                                !string.IsNullOrEmpty(interiorP)
                                && !toc.Success
                                && after == before
                                && Blob(toc, "目录域里面的编码"),
                                Dump(toc) + " before=" + before + " after=" + after);
                        }

                        if (Want("ul-table-toc"))
                        {
                            ToolResult created = s.CreateTableFromXml(TableArgs(id, WriteXml(GridXml("格内原文。", "", "备注", ""))));
                            string tableId = TableId(Body(s, id));
                            string code = string.IsNullOrEmpty(interiorS) ? interiorP : interiorS;
                            ToolResult wr = s.ProcessActions(id, Batch(
                                Replace(string.IsNullOrEmpty(code) ? "S_xxxxx" : code, "不应出现。"),
                                InsertTable(tableId, "不应出现。")));
                            string after = Body(s, id);
                            Ok(run, "ul-table-toc",
                                created.Success
                                && !string.IsNullOrEmpty(tableId)
                                && !wr.Success && WdSession.IntOf(wr, "executed_actions") == 0
                                && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0
                                && Blob(wr, "目录域里面的编码"),
                                Dump(wr) + " table=" + tableId + " body=" + Strip(after));
                        }
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }
        }

        private static void RunGeo(TestRun run, WdSession s)
        {
            if (!Want("ul-cell-write", "ul-cell-miss", "ul-cell-mix", "ul-cell-and-sent"))
            {
                return;
            }

            if (!Open(s, "正文留着。", out string path, out string id, out _))
            {
                FailMany(run, "铺表稿失败", "ul-cell-write", "ul-cell-miss", "ul-cell-mix", "ul-cell-and-sent");
                return;
            }

            try
            {
                s.CreateTableFromXml(TableArgs(id, WriteXml(GridXml("格内原文。", "", "备注", ""))));
                string body = Body(s, id);
                string sent = Sent(body, "正文留着");
                string empty = EmptyCell(body);
                string filled = FilledCell(body);

                if (Want("ul-cell-write"))
                {
                    ToolResult wr = s.ProcessActions(id, Replace(empty, "空格新字。"));
                    string after = Body(s, id);
                    Ok(run, "ul-cell-write",
                        wr.Success
                        && after.IndexOf("空格新字。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("格内原文。", StringComparison.Ordinal) >= 0,
                        Dump(wr) + " empty=" + empty + " body=" + Strip(after));
                    if (wr.Success)
                    {
                        s.ProcessActions(id, Replace(EmptyCell(after), ""));
                    }
                }

                body = Body(s, id);
                sent = Sent(body, "正文留着");
                empty = EmptyCell(body);

                if (Want("ul-cell-miss"))
                {
                    string tableId = TableId(body);
                    string miss = tableId + "_r9_c9";
                    ToolResult wr = s.ProcessActions(id, Batch(
                        Replace(miss, "不应出现。"),
                        Replace(sent, "正文已改。")));
                    string after = Body(s, id);
                    Ok(run, "ul-cell-miss",
                        !ItemOk(wr, 0) && ItemOk(wr, 1)
                        && after.IndexOf("正文已改。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("格内原文。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0,
                        Dump(wr) + " body=" + Strip(after));
                    sent = Sent(after, "正文已改");
                    if (ItemOk(wr, 1))
                    {
                        s.ProcessActions(id, Replace(sent, "正文留着。"));
                    }
                }

                body = Body(s, id);
                sent = Sent(body, "正文留着");
                empty = EmptyCell(body);

                if (Want("ul-cell-mix"))
                {
                    ToolResult wr = s.ProcessActions(id, Batch(
                        Replace(sent + "," + empty, "不应出现。"),
                        Replace(sent, "正文已改。")));
                    string after = Body(s, id);
                    Ok(run, "ul-cell-mix",
                        !ItemOk(wr, 0) && ItemOk(wr, 1)
                        && after.IndexOf("正文已改。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0,
                        Dump(wr) + " body=" + Strip(after));
                    if (ItemOk(wr, 1))
                    {
                        s.ProcessActions(id, Replace(Sent(after, "正文已改"), "正文留着。"));
                    }
                }

                body = Body(s, id);
                sent = Sent(body, "正文留着");
                empty = EmptyCell(body);

                if (Want("ul-cell-and-sent"))
                {
                    ToolResult wr = s.ProcessActions(id, Batch(
                        Replace(empty, "空格新字。"),
                        Replace(sent, "正文已改。")));
                    string after = Body(s, id);
                    Ok(run, "ul-cell-and-sent",
                        ItemOk(wr, 0) && ItemOk(wr, 1)
                        && after.IndexOf("空格新字。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("正文已改。", StringComparison.Ordinal) >= 0
                        && after.IndexOf("格内原文。", StringComparison.Ordinal) >= 0,
                        Dump(wr) + " empty=" + empty + " body=" + Strip(after));
                }
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunTable(TestRun run, WdSession s)
        {
            if (Want("ul-table-miss", "ul-table-gone", "ul-insert-table", "ul-del-table"))
            {
                if (!Open(s, "正文留着。", out string path, out string id, out _))
                {
                    FailMany(run, "铺表稿失败", "ul-table-miss", "ul-table-gone", "ul-insert-table", "ul-del-table");
                }
                else
                {
                    try
                    {
                        s.CreateTableFromXml(TableArgs(id, WriteXml(GridXml("格内原文。", "", "备注", ""))));
                        string body = Body(s, id);
                        string tableId = TableId(body);

                        if (Want("ul-table-miss"))
                        {
                            ToolResult wr = s.ProcessActions(id, Batch(
                                Replace("S_xxxxx", "不应出现。"),
                                InsertTable(tableId, "表后新段。")));
                            string after = Body(s, id);
                            Ok(run, "ul-table-miss",
                                !ItemOk(wr, 0) && ItemOk(wr, 1)
                                && after.IndexOf("表后新段。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("格内原文。", StringComparison.Ordinal) >= 0,
                                Dump(wr) + " body=" + Strip(after));
                        }

                        if (Want("ul-table-gone"))
                        {
                            body = Body(s, id);
                            string sent = Sent(body, "正文留着");
                            if (string.IsNullOrEmpty(sent))
                            {
                                sent = Sent(body, "正文");
                            }

                            ToolResult wr = s.ProcessActions(id, Batch(
                                Replace(sent, "正文已改。"),
                                InsertTable("T_xxxxx", "不应出现。")));
                            string after = Body(s, id);
                            Ok(run, "ul-table-gone",
                                ItemOk(wr, 0) && !ItemOk(wr, 1)
                                && after.IndexOf("正文已改。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0
                                && after.IndexOf("<table", StringComparison.Ordinal) >= 0,
                                Dump(wr) + " body=" + Strip(after));
                            if (ItemOk(wr, 0))
                            {
                                s.ProcessActions(id, Replace(Sent(after, "正文已改"), "正文留着。"));
                            }
                        }

                        if (Want("ul-insert-table"))
                        {
                            body = Body(s, id);
                            tableId = TableId(body);
                            ToolResult wr = s.ProcessActions(id, InsertTable(tableId, "表后新段乙。"));
                            string after = Body(s, id);
                            Ok(run, "ul-insert-table",
                                wr.Success
                                && after.IndexOf("表后新段乙。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("格内原文。", StringComparison.Ordinal) >= 0,
                                Dump(wr) + " body=" + Strip(after));
                        }

                        if (Want("ul-del-table"))
                        {
                            body = Body(s, id);
                            tableId = TableId(body);
                            ToolResult wr = s.DeleteTable(new Dictionary<string, object>
                            {
                                ["channel_id"] = id,
                                ["table_id"] = tableId
                            });
                            string after = Body(s, id);
                            Ok(run, "ul-del-table",
                                wr.Success
                                && after.IndexOf("<table", StringComparison.Ordinal) < 0
                                && after.IndexOf("正文留着。", StringComparison.Ordinal) >= 0,
                                Dump(wr) + " body=" + Strip(after));
                        }
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-comment-table"))
            {
                if (!Open(s, "正文留着。", out string path, out string id, out _))
                {
                    Fail(run, "ul-comment-table", "铺表稿失败");
                }
                else
                {
                    try
                    {
                        s.CreateTableFromXml(TableArgs(id, WriteXml(GridXml("格内原文。", "", "备注", ""))));
                        string tableId = TableId(Body(s, id));
                        ToolResult wr = s.ProcessComments(id, new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                ["type"] = "add",
                                ["text"] = "表后点。",
                                ["target"] = new Dictionary<string, object> { ["table"] = tableId }
                            }
                        });
                        bool pin = CommentAtTableEnd(s, path);
                        Ok(run, "ul-comment-table", wr.Success && pin, Dump(wr) + " pin=" + pin);
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-del-image"))
            {
                if (!Open(s, "正文留着。", out string path, out string id, out _))
                {
                    Fail(run, "ul-del-image", "铺稿失败");
                }
                else
                {
                    try
                    {
                        s.CreateTableFromXml(TableArgs(id, WriteXml(GridXml("格内原文。", "", "备注", ""))));
                        string png = Png();
                        ToolResult ins = s.InsertImage(new Dictionary<string, object>
                        {
                            ["channel_id"] = id,
                            ["path"] = png
                        });
                        string imageId = WdSession.Str(ins, "image_id");
                        if (string.IsNullOrEmpty(imageId))
                        {
                            imageId = ImageId(Body(s, id));
                        }

                        ToolResult wr = s.DeleteImage(new Dictionary<string, object>
                        {
                            ["channel_id"] = id,
                            ["image_id"] = imageId
                        });
                        string after = Body(s, id);
                        Ok(run, "ul-del-image",
                            ins.Success && wr.Success
                            && after.IndexOf("<image", StringComparison.Ordinal) < 0
                            && after.IndexOf("正文留着。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("<table", StringComparison.Ordinal) >= 0,
                            Dump(wr) + " image=" + imageId + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-chart-gone"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-chart-gone", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string yi = Sent(body, "乙句定位");
                        ToolResult wr = s.ProcessActions(id, Batch(
                            Replace(yi, "乙句已改。"),
                            InsertChart("C_xxxxx", "不应出现。")));
                        string after = Body(s, id);
                        Ok(run, "ul-chart-gone",
                            ItemOk(wr, 0) && !ItemOk(wr, 1)
                            && after.IndexOf("乙句已改。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("不应出现。", StringComparison.Ordinal) < 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }
        }

        private static void RunCallers(TestRun run, WdSession s)
        {
            if (Want("ul-replace-span", "ul-insert-end", "ul-comment-sent", "ul-text-slice", "ul-text-two", "ul-apply-part"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    FailMany(run, "铺稿失败",
                        "ul-replace-span", "ul-insert-end", "ul-comment-sent", "ul-text-slice",
                        "ul-text-two", "ul-apply-part");
                }
                else
                {
                    try
                    {
                        string jia = Sent(body, "甲句定位");
                        string yi = Sent(body, "乙句定位");

                        if (Want("ul-replace-span"))
                        {
                            ToolResult wr = s.ProcessActions(id, Replace(jia, "甲句换成。"));
                            string after = Body(s, id);
                            Ok(run, "ul-replace-span",
                                wr.Success
                                && after.IndexOf("甲句换成。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("甲句定位。", StringComparison.Ordinal) < 0,
                                Dump(wr) + " body=" + Strip(after));
                            if (wr.Success)
                            {
                                s.ProcessActions(id, Replace(Sent(after, "甲句换成"), "甲句定位。"));
                            }
                        }

                        body = Body(s, id);
                        jia = Sent(body, "甲句定位");
                        yi = Sent(body, "乙句定位");

                        if (Want("ul-insert-end"))
                        {
                            ToolResult wr = s.ProcessActions(id, InsertAfter(jia, "插在甲后。"));
                            string after = Body(s, id);
                            int a = after.IndexOf("甲句定位。", StringComparison.Ordinal);
                            int b = after.IndexOf("插在甲后。", StringComparison.Ordinal);
                            int c = after.IndexOf("乙句定位。", StringComparison.Ordinal);
                            Ok(run, "ul-insert-end",
                                wr.Success && a >= 0 && b > a && c > b,
                                Dump(wr) + " body=" + Strip(after));
                        }

                        body = Body(s, id);
                        jia = Sent(body, "甲句定位");
                        yi = Sent(body, "乙句定位");

                        if (Want("ul-comment-sent"))
                        {
                            ToolResult wr = s.ProcessComments(id, new List<Dictionary<string, object>>
                            {
                                new Dictionary<string, object>
                                {
                                    ["type"] = "add",
                                    ["text"] = "句上。",
                                    ["target"] = new Dictionary<string, object> { ["codes"] = jia }
                                }
                            });
                            bool covers = CommentCovers(s, path, "甲句定位");
                            Ok(run, "ul-comment-sent", wr.Success && covers, Dump(wr) + " covers=" + covers);
                        }

                        if (Want("ul-text-slice"))
                        {
                            body = Body(s, id);
                            jia = Sent(body, "甲句定位");
                            var action = One("replace", new Dictionary<string, object>
                            {
                                ["original"] = new Dictionary<string, object> { ["codes"] = jia, ["text"] = "定位" },
                                ["new"] = "已改",
                                ["format"] = DefaultFormat()
                            });
                            ToolResult wr = s.ProcessActions(id, action);
                            string after = Body(s, id);
                            Ok(run, "ul-text-slice",
                                wr.Success
                                && after.IndexOf("甲句已改。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("甲句定位。", StringComparison.Ordinal) < 0,
                                Dump(wr) + " body=" + Strip(after));
                            if (wr.Success)
                            {
                                s.ProcessActions(id, Replace(Sent(after, "甲句已改"), "甲句定位。"));
                            }
                        }

                        body = Body(s, id);
                        jia = Sent(body, "甲句定位");
                        yi = Sent(body, "乙句定位");

                        if (Want("ul-text-two"))
                        {
                            var action = One("replace", new Dictionary<string, object>
                            {
                                ["original"] = new Dictionary<string, object>
                                {
                                    ["codes"] = jia + "," + yi,
                                    ["text"] = "定位"
                                },
                                ["new"] = "不应出现。",
                                ["format"] = DefaultFormat()
                            });
                            ToolResult wr = s.ProcessActions(id, Batch(action, Replace(yi, "乙句已改。")));
                            string after = Body(s, id);
                            Ok(run, "ul-text-two",
                                !ItemOk(wr, 0) && ItemOk(wr, 1)
                                && Blob(wr, "恰好一个")
                                && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0
                                && after.IndexOf("乙句已改。", StringComparison.Ordinal) >= 0,
                                Dump(wr) + " body=" + Strip(after));
                            if (ItemOk(wr, 1))
                            {
                                s.ProcessActions(id, Replace(Sent(after, "乙句已改"), "乙句定位。"));
                            }
                        }

                        body = Body(s, id);
                        jia = Sent(body, "甲句定位");
                        yi = Sent(body, "乙句定位");

                        if (Want("ul-apply-part"))
                        {
                            ToolResult applied = Apply(s, id, jia + ",S_xxxxx", "heading_1");
                            bool jiaStyled = StyleHas(s, path, "甲句定位。", HeadingNames);
                            bool yiStyled = StyleHas(s, path, "乙句定位。", HeadingNames);
                            Ok(run, "ul-apply-part",
                                !applied.Success && !jiaStyled && !yiStyled,
                                Dump(applied) + " jia=" + jiaStyled + " yi=" + yiStyled);
                        }
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-apply-ok"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-apply-ok", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string jia = Sent(body, "甲句定位");
                        string yi = Sent(body, "乙句定位");
                        ToolResult applied = Apply(s, id, jia + "," + yi, "heading_1");
                        bool jiaStyled = StyleHas(s, path, "甲句定位。", HeadingNames);
                        bool yiStyled = StyleHas(s, path, "乙句定位。", HeadingNames);
                        Ok(run, "ul-apply-ok",
                            applied.Success && jiaStyled && yiStyled,
                            Dump(applied) + " jia=" + jiaStyled + " yi=" + yiStyled);
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-para-span"))
            {
                if (!Open(s, "甲句定位。乙句定位。", out string path, out string id, out _))
                {
                    Fail(run, "ul-para-span", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string para = Para(ParagraphBody(s, id), "甲句定位");
                        var action = One("replace", new Dictionary<string, object>
                        {
                            ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = para },
                            ["new"] = "整段换成。",
                            ["format"] = DefaultFormat()
                        });
                        ToolResult wr = s.ProcessActions(id, action);
                        string after = Body(s, id);
                        Ok(run, "ul-para-span",
                            wr.Success
                            && after.IndexOf("整段换成。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) < 0
                            && after.IndexOf("乙句定位。", StringComparison.Ordinal) < 0,
                            Dump(wr) + " para=" + para + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }
        }

        private static void RunOut(TestRun run, WdSession s)
        {
            if (Want("ul-blank"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out _))
                {
                    Fail(run, "ul-blank", "铺稿失败");
                }
                else
                {
                    try
                    {
                        ToolResult wr = s.InsertBlankPage(new Dictionary<string, object>
                        {
                            ["channel_id"] = id,
                            ["position"] = "start"
                        });
                        string after = Body(s, id);
                        Ok(run, "ul-blank",
                            wr.Success
                            && after.IndexOf("甲句定位。", StringComparison.Ordinal) >= 0
                            && after.IndexOf("乙句定位。", StringComparison.Ordinal) >= 0,
                            Dump(wr) + " body=" + Strip(after));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-del-comment"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out string body))
                {
                    Fail(run, "ul-del-comment", "铺稿失败");
                }
                else
                {
                    try
                    {
                        string jia = Sent(body, "甲句定位");
                        s.ProcessComments(id, new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                ["type"] = "add",
                                ["text"] = "句上。",
                                ["target"] = new Dictionary<string, object> { ["codes"] = jia }
                            }
                        });
                        ToolResult wr = s.ProcessComments(id, new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { ["type"] = "delete", ["index"] = 1 }
                        });
                        Ok(run, "ul-del-comment", wr.Success && CommentCount(s, path) == 0, Dump(wr));
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }

            if (Want("ul-comment-cursor"))
            {
                if (!Open(s, "甲句定位。\r乙句定位。", out string path, out string id, out _))
                {
                    Fail(run, "ul-comment-cursor", "铺稿失败");
                }
                else
                {
                    try
                    {
                        SelectNeedle(s, path, "甲句定位。");
                        ToolResult wr = s.ProcessComments(id, new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object> { ["type"] = "add", ["text"] = "选区。" }
                        });
                        bool covers = CommentCovers(s, path, "甲句定位");
                        Ok(run, "ul-comment-cursor", wr.Success && covers, Dump(wr) + " covers=" + covers);
                    }
                    finally
                    {
                        s.CloseDocByPath(path);
                    }
                }
            }
        }

        private static bool Open(WdSession s, string plain, out string path, out string id, out string body)
        {
            path = s.NewDocxPath();
            id = null;
            body = "";
            ToolResult opened = s.Open(path, true, s.App);
            if (opened == null || !opened.Success)
            {
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            if (string.IsNullOrEmpty(id) || !s.TypePlain(path, plain))
            {
                return false;
            }

            body = Body(s, id);
            return !string.IsNullOrEmpty(body);
        }

        private static string Body(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "sentence");
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string ParagraphBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "paragraph");
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string Sent(string body, string text)
        {
            if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(text))
            {
                return "";
            }

            int at = body.IndexOf(text, StringComparison.Ordinal);
            if (at < 0)
            {
                return "";
            }

            int mark = body.LastIndexOf("[S_", at, StringComparison.Ordinal);
            if (mark < 0)
            {
                return "";
            }

            Match m = SentMark.Match(body, mark);
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string Para(string body, string text)
        {
            if (string.IsNullOrEmpty(body) || string.IsNullOrEmpty(text))
            {
                return "";
            }

            int at = body.IndexOf(text, StringComparison.Ordinal);
            if (at < 0)
            {
                return "";
            }

            int mark = body.LastIndexOf("[P_", at, StringComparison.Ordinal);
            if (mark < 0)
            {
                return "";
            }

            Match m = ParaMark.Match(body, mark);
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string TableId(string body)
        {
            Match m = TableMark.Match(body ?? "");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string ImageId(string body)
        {
            Match m = ImageMark.Match(body ?? "");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string EmptyCell(string body)
        {
            foreach (Match m in CellMark.Matches(body ?? ""))
            {
                string code = m.Groups[1].Value;
                int start = m.Index + m.Length;
                int end = body.IndexOf("[T_", start, StringComparison.Ordinal);
                string inner = end < 0 ? body.Substring(start) : body.Substring(start, end - start);
                if (inner.IndexOf("[EMPTY]", StringComparison.Ordinal) >= 0 || Strip(inner).Length == 0)
                {
                    return code;
                }
            }

            return "";
        }

        private static string FilledCell(string body)
        {
            foreach (Match m in CellMark.Matches(body ?? ""))
            {
                string code = m.Groups[1].Value;
                if (code != EmptyCell(body))
                {
                    return code;
                }
            }

            return "";
        }

        private static List<Dictionary<string, object>> One(string type, Dictionary<string, object> detail)
        {
            return new List<Dictionary<string, object>>
            {
                new Dictionary<string, object> { ["type"] = type, ["detail"] = detail }
            };
        }

        private static List<Dictionary<string, object>> Batch(params List<Dictionary<string, object>>[] parts)
        {
            var list = new List<Dictionary<string, object>>();
            if (parts == null)
            {
                return list;
            }

            foreach (List<Dictionary<string, object>> part in parts)
            {
                if (part != null)
                {
                    list.AddRange(part);
                }
            }

            return list;
        }

        private static List<Dictionary<string, object>> Replace(string codes, string neu)
        {
            return One("replace", new Dictionary<string, object>
            {
                ["original"] = new Dictionary<string, object> { ["codes"] = codes },
                ["new"] = neu,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertAfter(string codes, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["codes"] = codes },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertTable(string tableId, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["table"] = tableId },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static List<Dictionary<string, object>> InsertChart(string chartId, string insert)
        {
            return One("insert", new Dictionary<string, object>
            {
                ["target"] = new Dictionary<string, object> { ["chart"] = chartId },
                ["insert"] = insert,
                ["format"] = DefaultFormat()
            });
        }

        private static Dictionary<string, object> DefaultFormat()
        {
            return new Dictionary<string, object>
            {
                ["char_format"] = new Dictionary<string, object>
                {
                    ["font_name"] = "宋体",
                    ["font_size"] = 12,
                    ["is_bold"] = false
                }
            };
        }

        private static Dictionary<string, object> TableArgs(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["path"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static string GridXml(string a, string b, string c, string d)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "<tr><td>" + c + "</td><td>" + d + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-ul-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
        }

        private static string Png()
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-ul-" + Guid.NewGuid().ToString("N") + ".png");
            using (var bmp = new Bitmap(96, 48))
            using (var g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.FromArgb(0x44, 0x72, 0xC4));
                bmp.Save(path, ImageFormat.Png);
            }

            return path;
        }

        private static ToolResult Toc(string channel, string action, string tocId, string at)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channel))
            {
                args["channel_id"] = channel;
            }

            args["action"] = action;
            if (!string.IsNullOrEmpty(tocId))
            {
                args["toc_id"] = tocId;
            }

            if (!string.IsNullOrEmpty(at))
            {
                args["at"] = at;
            }

            return DesktopToolPipe.Call("F_toc", args);
        }

        private static ToolResult Apply(WdSession s, string id, string codes, string structure)
        {
            return s.DocumentStructure(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["action"] = "apply",
                ["target_codes"] = codes,
                ["structure"] = structure
            });
        }

        private static bool StyleHeading(WdSession s, string path)
        {
            if (s.IsWps)
            {
                return s.TryWpsDoc(path, doc =>
                {
                    object paragraph = WdSession.ComGet(doc, "Paragraphs");
                    object first = WdSession.ComCall(paragraph, "Item", 1);
                    object styles = WdSession.ComGet(doc, "Styles");
                    object style = null;
                    foreach (string name in HeadingNames)
                    {
                        try
                        {
                            style = WdSession.ComCall(styles, "Item", name);
                            if (style != null)
                            {
                                break;
                            }
                        }
                        catch (Exception)
                        {
                            style = null;
                        }
                    }

                    if (first == null || style == null)
                    {
                        return false;
                    }

                    WdSession.ComSet(first, "Style", style);
                    return true;
                });
            }

            return s.OnWordDoc(path, doc =>
            {
                if (doc.Paragraphs.Count < 1)
                {
                    return false;
                }

                foreach (string name in HeadingNames)
                {
                    try
                    {
                        object style = doc.Styles[name];
                        doc.Paragraphs[1].set_Style(ref style);
                        return true;
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
        }

        private static object CommentScope(object comment)
        {
            try
            {
                object scope = WdSession.ComGet(comment, "Scope");
                if (scope != null)
                {
                    return scope;
                }
            }
            catch (Exception)
            {
            }

            return WdSession.ComGet(comment, "Range");
        }

        private static Word.Range CommentScopeRange(Word.Comment comment)
        {
            try
            {
                if (comment.Scope != null)
                {
                    return comment.Scope;
                }
            }
            catch (Exception)
            {
            }

            return comment.Range;
        }

        private static string CommentScopeText(Word.Comment comment)
        {
            Word.Range range = CommentScopeRange(comment);
            return range == null ? "" : (range.Text ?? "");
        }

        private static bool StyleHas(WdSession s, string path, string needle, string[] names)
        {
            string found = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    found = ParaStyleName(doc, needle);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range content = doc.Content;
                    string text = content.Text ?? "";
                    int at = text.IndexOf(needle, StringComparison.Ordinal);
                    if (at < 0)
                    {
                        return false;
                    }

                    Word.Range range = doc.Range(content.Start + at, content.Start + at + Math.Min(needle.Length, 1));
                    found = range.get_Style() is Word.Style style ? (style.NameLocal ?? "") : "";
                    return found.Length > 0;
                });
            if (!ok || string.IsNullOrEmpty(found))
            {
                return false;
            }

            foreach (string name in names)
            {
                if (found.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string ParaStyleName(object doc, string needle)
        {
            try
            {
                object content = WdSession.ComGet(doc, "Content");
                string text = Convert.ToString(WdSession.ComGet(content, "Text")) ?? "";
                int at = text.IndexOf(needle, StringComparison.Ordinal);
                if (at < 0)
                {
                    return "";
                }

                int start = Convert.ToInt32(WdSession.ComGet(content, "Start")) + at;
                object range = WdSession.ComCall(doc, "Range", start, start + 1);
                object style = WdSession.ComGet(range, "Style");
                return Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool CommentCovers(WdSession s, string path, string needle)
        {
            bool ok = false;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object comments = WdSession.ComGet(doc, "Comments");
                    int n = Convert.ToInt32(WdSession.ComGet(comments, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object comment = WdSession.ComCall(comments, "Item", i);
                        object range = CommentScope(comment);
                        string text = Convert.ToString(WdSession.ComGet(range, "Text")) ?? "";
                        if (text.IndexOf(needle, StringComparison.Ordinal) >= 0)
                        {
                            ok = true;
                            return true;
                        }
                    }

                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    for (int i = 1; i <= doc.Comments.Count; i++)
                    {
                        string text = CommentScopeText(doc.Comments[i]);
                        if (text.IndexOf(needle, StringComparison.Ordinal) >= 0)
                        {
                            ok = true;
                            return true;
                        }
                    }

                    return true;
                });
            }

            return ok;
        }

        private static bool CommentAtTableEnd(WdSession s, string path)
        {
            bool ok = false;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object tables = WdSession.ComGet(doc, "Tables");
                    if (Convert.ToInt32(WdSession.ComGet(tables, "Count")) < 1)
                    {
                        return false;
                    }

                    object table = WdSession.ComCall(tables, "Item", 1);
                    object tableRange = WdSession.ComGet(table, "Range");
                    int tableEnd = Convert.ToInt32(WdSession.ComGet(tableRange, "End"));
                    object comments = WdSession.ComGet(doc, "Comments");
                    int n = Convert.ToInt32(WdSession.ComGet(comments, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object comment = WdSession.ComCall(comments, "Item", i);
                        object range = CommentScope(comment);
                        int start = Convert.ToInt32(WdSession.ComGet(range, "Start"));
                        int end = Convert.ToInt32(WdSession.ComGet(range, "End"));
                        if (start == tableEnd && end == start)
                        {
                            ok = true;
                            return true;
                        }
                    }

                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    if (doc.Tables.Count < 1 || doc.Comments.Count < 1)
                    {
                        return false;
                    }

                    int tableEnd = doc.Tables[1].Range.End;
                    for (int i = 1; i <= doc.Comments.Count; i++)
                    {
                        Word.Range range = CommentScopeRange(doc.Comments[i]);
                        if (range.Start == tableEnd && range.End == range.Start)
                        {
                            ok = true;
                            return true;
                        }
                    }

                    return true;
                });
            }

            return ok;
        }

        private static int CommentCount(WdSession s, string path)
        {
            int n = -1;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    n = Convert.ToInt32(WdSession.ComGet(WdSession.ComGet(doc, "Comments"), "Count"));
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    n = doc.Comments.Count;
                    return true;
                });
            }

            return n;
        }

        private static void SelectNeedle(WdSession s, string path, string needle)
        {
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object content = WdSession.ComGet(doc, "Content");
                    string text = Convert.ToString(WdSession.ComGet(content, "Text")) ?? "";
                    int at = text.IndexOf(needle, StringComparison.Ordinal);
                    if (at < 0)
                    {
                        return false;
                    }

                    int start = Convert.ToInt32(WdSession.ComGet(content, "Start")) + at;
                    object app = WdSession.ComGet(doc, "Application");
                    object sel = WdSession.ComGet(app, "Selection");
                    WdSession.ComCall(sel, "SetRange", start, start + needle.Length);
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    Word.Range content = doc.Content;
                    string text = content.Text ?? "";
                    int at = text.IndexOf(needle, StringComparison.Ordinal);
                    if (at < 0)
                    {
                        return false;
                    }

                    doc.Application.Selection.SetRange(content.Start + at, content.Start + at + needle.Length);
                    return true;
                });
            }
        }

        private static bool ItemOk(ToolResult r, int index)
        {
            JArray arr = Results(r);
            if (arr == null || index < 0 || index >= arr.Count)
            {
                return false;
            }

            JToken token = arr[index];
            if (token is JObject obj)
            {
                return obj.Value<bool?>("success") == true;
            }

            return false;
        }

        private static JArray Results(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return null;
            }

            try
            {
                JObject jo = r.Data is Dictionary<string, object>
                    ? JObject.FromObject(r.Data)
                    : JObject.Parse(JsonConvert.SerializeObject(r.Data));
                return jo["results"] as JArray;
            }
            catch
            {
                return null;
            }
        }

        private static bool Blob(ToolResult r, string needle)
        {
            if (r != null && !string.IsNullOrEmpty(r.Error) && r.Error.IndexOf(needle, StringComparison.Ordinal) >= 0)
            {
                return true;
            }

            if (r == null || r.Data == null)
            {
                return false;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data).IndexOf(needle, StringComparison.Ordinal) >= 0;
            }
            catch
            {
                return Convert.ToString(r.Data).IndexOf(needle, StringComparison.Ordinal) >= 0;
            }
        }

        private static string Dump(ToolResult r)
        {
            if (r == null)
            {
                return "没有回包";
            }

            if (!string.IsNullOrEmpty(r.Error))
            {
                return r.Error;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data);
            }
            catch
            {
                return Convert.ToString(r.Data) ?? "";
            }
        }

        private static string Strip(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "";
            }

            string t = Regex.Replace(body, @"\[[^\]]+\]", "");
            t = Regex.Replace(t, @"<[^>]+>", "");
            return t.Length > 180 ? t.Substring(0, 180) : t;
        }

        private static int Count(string text, string needle)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle))
            {
                return 0;
            }

            int n = 0;
            int from = 0;
            while (from <= text.Length - needle.Length)
            {
                int i = text.IndexOf(needle, from, StringComparison.Ordinal);
                if (i < 0)
                {
                    break;
                }

                n++;
                from = i + needle.Length;
            }

            return n;
        }

        private static bool Want(params string[] ids)
        {
            if (Only == null)
            {
                return true;
            }

            foreach (string id in ids)
            {
                if (Only.Contains(id))
                {
                    return true;
                }
            }

            return false;
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            Seen.Add(id);
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
            Seen.Add(id);
            if (Want(id))
            {
                run.CaseFail(id, reason);
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                Fail(run, id, reason);
            }
        }
    }
}
