using System;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>F_toc 的增删改查。对照 docs/测试/word测试/目录工具。不进 tc- 90。</summary>
    internal static class TuCases
    {
        public const int Target = 157;

        private static readonly Regex TocToken = new Regex(
            @"\[目录域 (ewC_[A-Za-z0-9]{8,32})\]",
            RegexOptions.Compiled);

        private static readonly Regex ParaCode = new Regex(
            @"P_[A-Za-z0-9]+",
            RegexOptions.Compiled);

        private static HashSet<string> Only;
        private static readonly HashSet<string> Seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public static void Run(TestRun run, string host, string filter)
        {
            Seen.Clear();
            Only = null;
            if (!string.IsNullOrWhiteSpace(filter)
                && !filter.Equals("tu", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("tu-", StringComparison.OrdinalIgnoreCase))
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

            var session = new WdSession(host);
            RunContract(run, session);
            RunContent(run, session);
            RunFormat(run, session);
            RunActions(run, session);
            RunKeys(run, session);
            foreach (string id in AllIds())
            {
                if (!Want(id) || Seen.Contains(id))
                {
                    continue;
                }

                Fail(run, id, "测试方法没有执行这条");
            }
        }

        private static void RunContract(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-x-action", "tu-x-no-action", "tu-x-no-id", "tu-x-bad-id", "tu-x-no-xml",
                "tu-x-root", "tu-x-levels", "tu-x-style", "tu-x-sep", "tu-x-leader", "tu-x-font", "tu-x-range"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel))
            {
                FailMany(run, "打不开临时稿", "tu-x-action", "tu-x-no-action", "tu-x-no-id", "tu-x-bad-id",
                    "tu-x-no-xml", "tu-x-root", "tu-x-levels", "tu-x-style", "tu-x-sep", "tu-x-leader",
                    "tu-x-font", "tu-x-range");
                return;
            }

            try
            {
                CheckFail(run, session, channel, "tu-x-action", "getx", null, null, null, "action 只能是");
                CheckFail(run, session, channel, "tu-x-no-action", null, null, null, null, "未提供 action");
                CheckFail(run, session, channel, "tu-x-no-id", "get", null, null, null, "未提供 toc_id");
                CheckFail(run, session, channel, "tu-x-bad-id", "get", "TOC1", null, null, "ewC_");
                CheckFail(run, session, channel, "tu-x-no-xml", "update", "ewC_abcdefgh", null, null, "须提供 xml");
                CheckFail(run, session, channel, "tu-x-root", "update", "ewC_abcdefgh", "<doc/>", null, "toc 为根");
                CheckFail(run, session, channel, "tu-x-levels", "update", "ewC_abcdefgh",
                    "<toc><content><levels>9-1</levels></content></toc>", null, "起点不大于终点");
                CheckFail(run, session, channel, "tu-x-style", "update", "ewC_abcdefgh",
                    "<toc><content><styles><style name=\"甲,乙\" level=\"1\"/></styles></content></toc>", null, "逗号");
                CheckFail(run, session, channel, "tu-x-sep", "update", "ewC_abcdefgh",
                    "<toc><format><separator>123456</separator></format></toc>", null, "5");
                CheckFail(run, session, channel, "tu-x-leader", "update", "ewC_abcdefgh",
                    "<toc><format><leader>wave</leader></format></toc>", null, "前导符");
                CheckFail(run, session, channel, "tu-x-font", "update", "ewC_abcdefgh",
                    "<toc><format><level n=\"1\"><font zz_font=\"1\"/></level></format></toc>", null, "不认识的字符格式");
                CheckFail(run, session, channel, "tu-x-range", "update", "ewC_abcdefgh",
                    "<toc><content><range from=\"P_a\"/></content></toc>", null, "起止要成对");
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunContent(TestRun run, WdSession session)
        {
            string[] ids =
            {
                "tu-levels-default", "tu-levels-11", "tu-levels-19", "tu-levels-99",
                "tu-styles-one", "tu-styles-clear", "tu-styles-keep",
                "tu-range-pair", "tu-range-clear", "tu-range-half"
            };
            if (!Need(run, ids))
            {
                return;
            }

            if (!Open(session, "第一章标题\r第二章标题\r正文停留。", out string path, out string channel)
                || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", ids);
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null, null, "start");
                string id = WdSession.Str(inserted, "toc_id");
                string xml = inserted.Success
                    ? WdSession.Str(Toc(channel, "get", id, null, null), "xml")
                    : inserted.Error;
                Ok(run, "tu-levels-default",
                    inserted.Success && xml.Contains("<levels>1-3</levels>") && !xml.Contains("<hyperlink") && !xml.Contains("<styles"),
                    xml);

                OkXml(run, session, channel, id, "tu-levels-11",
                    "<toc><content><levels>1-1</levels></content></toc>", "<levels>1-1</levels>");
                OkXml(run, session, channel, id, "tu-levels-19",
                    "<toc><content><levels>1-9</levels></content></toc>", "<levels>1-9</levels>");
                OkXml(run, session, channel, id, "tu-levels-99",
                    "<toc><content><levels>9-9</levels></content></toc>", "<levels>9-9</levels>");

                string styled = "<toc><content><levels>1-3</levels><styles><style name=\"标题 1\" level=\"2\"/></styles></content></toc>";
                OkXml(run, session, channel, id, "tu-styles-one", styled, "name=\"标题 1\"");
                OkXml(run, session, channel, id, "tu-styles-keep",
                    "<toc><content><levels>1-2</levels></content></toc>", "name=\"标题 1\"");
                string cleared = GetXml(session, channel, id,
                    "<toc><content><styles></styles></content></toc>");
                Ok(run, "tu-styles-clear", cleared.IndexOf("<styles", StringComparison.Ordinal) < 0, cleared);

                List<string> codes = ParaCodes(session, channel);
                if (codes.Count < 2)
                {
                    Fail(run, "tu-range-pair", "段落码不够：" + string.Join(",", codes));
                    Fail(run, "tu-range-clear", "段落码不够");
                }
                else
                {
                    string rangeXml = "<toc><content><range from=\"" + codes[0] + "\" to=\"" + codes[1] + "\"/></content></toc>";
                    ToolResult ranged = Toc(channel, "update", id, rangeXml, null);
                    Ok(run, "tu-range-pair", ranged != null && ranged.Success, ranged == null ? "没有回包" : (ranged.Error ?? ""));
                    string gone = GetXml(session, channel, id, "<toc><content><range/></content></toc>");
                    Ok(run, "tu-range-clear", gone.IndexOf("<range", StringComparison.Ordinal) < 0, gone);
                }

                CheckFail(run, session, channel, "tu-range-half", "update", id,
                    "<toc><content><range from=\"P_a\"/></content></toc>", null, "起止要成对");
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunFormat(TestRun run, WdSession session)
        {
            string[] ids =
            {
                "tu-page-off", "tu-page-49", "tu-align-after", "tu-leader-none", "tu-leader-dash",
                "tu-leader-line", "tu-leader-heavy", "tu-leader-middle", "tu-hyper-off", "tu-web-off",
                "tu-chapter", "tu-keep-tabs", "tu-keep-breaks", "tu-sep-five"
            };
            if (!Need(run, ids))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel) || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", ids);
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null, null, null);
                string id = WdSession.Str(inserted, "toc_id");
                if (!inserted.Success || id.Length == 0)
                {
                    FailMany(run, inserted.Error ?? "插入失败", ids);
                    return;
                }

                OkXml(run, session, channel, id, "tu-page-off",
                    "<toc><format><pageNumbers show=\"false\"/></format></toc>", "show=\"false\"");
                OkXml(run, session, channel, id, "tu-page-49",
                    "<toc><format><pageNumbers levels=\"4-9\" show=\"false\"/></format></toc>", "levels=\"4-9\"");
                OkXml(run, session, channel, id, "tu-align-after",
                    "<toc><format><align>after</align></format></toc>", "after");
                OkLeader(run, session, channel, id, "tu-leader-none", "none", "wdTabLeaderSpaces");
                OkLeader(run, session, channel, id, "tu-leader-dash", "dash", "wdTabLeaderDashes");
                OkLeader(run, session, channel, id, "tu-leader-line", "line", "wdTabLeaderLines");
                OkLeader(run, session, channel, id, "tu-leader-heavy", "heavy", "wdTabLeaderHeavy");
                OkLeader(run, session, channel, id, "tu-leader-middle", "middleDot", "wdTabLeaderMiddleDot");
                OkXml(run, session, channel, id, "tu-hyper-off",
                    "<toc><format><hyperlink>false</hyperlink></format></toc>", "<hyperlink>false</hyperlink>");
                OkXml(run, session, channel, id, "tu-web-off",
                    "<toc><format><webHidePageNumbers>false</webHidePageNumbers></format></toc>",
                    "<webHidePageNumbers>false</webHidePageNumbers>");
                OkXml(run, session, channel, id, "tu-chapter",
                    "<toc><format><chapter seq=\"chapter\" separator=\"-\"/></format></toc>", "seq=\"chapter\"");
                OkXml(run, session, channel, id, "tu-keep-tabs",
                    "<toc><format><keepTabs>true</keepTabs></format></toc>", "<keepTabs>true</keepTabs>");
                OkXml(run, session, channel, id, "tu-keep-breaks",
                    "<toc><format><keepBreaks>true</keepBreaks></format></toc>", "<keepBreaks>true</keepBreaks>");
                OkXml(run, session, channel, id, "tu-sep-five",
                    "<toc><format><align>after</align><separator>.....</separator></format></toc>", ".....");
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunActions(TestRun run, WdSession session)
        {
            string[] ids =
            {
                "tu-title-on", "tu-title-off", "tu-title-level", "tu-get-one", "tu-get-text", "tu-get-omit",
                "tu-insert-at", "tu-insert-interior", "tu-update-part", "tu-format-local", "tu-update-gallery",
                "tu-delete-wrap", "tu-delete-plain", "tu-reset", "tu-style-shared", "tu-empty-ok"
            };
            if (!Need(run, ids))
            {
                return;
            }

            RunTitles(run, session);
            RunReads(run, session);
            RunPlace(run, session);
            RunTwo(run, session);
            RunGallery(run, session);
            RunDeletePlain(run, session);
            RunEmpty(run, session);
        }

        private static void RunTitles(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-title-on", "tu-title-off", "tu-title-level"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel))
            {
                FailMany(run, "打不开临时稿", "tu-title-on", "tu-title-off", "tu-title-level");
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null,
                    "<toc><format><title text=\"目录\"/></format></toc>", null);
                string id = WdSession.Str(inserted, "toc_id");
                string xml = inserted.Success
                    ? WdSession.Str(Toc(channel, "get", id, null, null), "xml")
                    : inserted.Error;
                Ok(run, "tu-title-on", inserted.Success && xml.Contains("text=\"目录\""), xml);
                string off = GetXml(session, channel, id, "<toc><format><title/></format></toc>");
                Ok(run, "tu-title-off", off.IndexOf("<title", StringComparison.Ordinal) < 0, off);
                string level = GetXml(session, channel, id,
                    "<toc><format><title text=\"目录\"/><level n=\"title\"><font font_size=\"22\"/></level></format></toc>");
                Ok(run, "tu-title-level", level.Contains("n=\"title\"") && level.Contains("22"), level);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunReads(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-get-text", "tu-get-omit", "tu-update-part"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel) || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", "tu-get-text", "tu-get-omit", "tu-update-part");
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null, null, null);
                string id = WdSession.Str(inserted, "toc_id");
                ToolResult got = Toc(channel, "get", id, null, null);
                string xml = WdSession.Str(got, "xml");
                string lines = LineText(got);
                Ok(run, "tu-get-omit",
                    got.Success && xml.Contains("<levels>1-3</levels>")
                    && xml.IndexOf("<hyperlink", StringComparison.Ordinal) < 0
                    && xml.IndexOf("<align", StringComparison.Ordinal) < 0
                    && xml.IndexOf("<leader", StringComparison.Ordinal) < 0,
                    xml);
                Ok(run, "tu-get-text",
                    got.Success && (lines.Contains("第一章") || lines.Contains("没有找到") || WdSession.Bool(got, "empty") || lines.Length > 0),
                    lines.Length == 0 ? xml : lines);

                Toc(channel, "update", id,
                    "<toc><format><level n=\"1\"><font font_size=\"16\"/></level></format></toc>", null);
                Toc(channel, "update", id,
                    "<toc><content><styles><style name=\"标题 1\" level=\"2\"/></styles></content></toc>", null);
                string part = GetXml(session, channel, id, "<toc><content><levels>1-2</levels></content></toc>");
                Ok(run, "tu-update-part",
                    part.Contains("<levels>1-2</levels>") && part.Contains("标题 1") && part.Contains("16"),
                    part);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunPlace(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-insert-at", "tu-insert-interior"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel) || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", "tu-insert-at", "tu-insert-interior");
                return;
            }

            try
            {
                Toc(channel, "insert", null, null, "start");
                List<string> codes = ParaCodes(session, channel);
                string body = codes.Count == 0 ? "" : codes[codes.Count - 1];
                ToolResult at = Toc(channel, "insert", null, null, body);
                string display = Paragraph(session, channel);
                Ok(run, "tu-insert-at",
                    at.Success && body.StartsWith("P_", StringComparison.Ordinal) && display.Contains(WdSession.Str(at, "toc_id")),
                    at.Success ? display : at.Error);

                string interior = FirstInteriorParagraph(session, path);
                if (interior.Length == 0)
                {
                    Fail(run, "tu-insert-interior", "没有读到目录里面的段码");
                }
                else
                {
                    ToolResult bad = Toc(channel, "insert", null, null, interior);
                    Ok(run, "tu-insert-interior",
                        !bad.Success && (bad.Error ?? "").Contains("目录域里面的编码"),
                        bad.Error ?? "");
                }
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunTwo(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-get-one", "tu-format-local", "tu-reset", "tu-style-shared"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r第二章标题\r正文停留。", out string path, out string channel) || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", "tu-get-one", "tu-format-local", "tu-reset", "tu-style-shared");
                return;
            }

            try
            {
                ToolResult first = Toc(channel, "insert", null, null, "start");
                ToolResult second = Toc(channel, "insert", null, null, "start");
                string a = WdSession.Str(first, "toc_id");
                string b = WdSession.Str(second, "toc_id");
                ToolResult got = Toc(channel, "get", a, null, null);
                string xml = WdSession.Str(got, "xml");
                Ok(run, "tu-get-one",
                    got.Success && a.Length > 0 && b.Length > 0 && a != b && !xml.Contains(b),
                    xml);

                string sized = GetXml(session, channel, a,
                    "<toc><format><level n=\"1\"><font font_size=\"16\"/></level></format></toc>");
                string other = WdSession.Str(Toc(channel, "get", b, null, null), "xml");
                bool styleUntouched = StyleSizeIsNot(session, path, 16f);
                Ok(run, "tu-format-local",
                    sized.Contains("16") && other.IndexOf("font_size=\"16\"", StringComparison.Ordinal) < 0 && styleUntouched,
                    "本份=" + sized + " 另一份=" + other);

                string reset = GetXml(session, channel, a, "<toc><format reset=\"true\"/></toc>");
                Ok(run, "tu-reset",
                    reset.Contains("<levels>") && reset.IndexOf("font_size=\"16\"", StringComparison.Ordinal) < 0,
                    reset);

                bool styled = SetTocStyleSize(session, path, 20f);
                string shared = WdSession.Str(Toc(channel, "get", b, null, null), "xml");
                Ok(run, "tu-style-shared",
                    styled && shared.IndexOf("font_size=\"16\"", StringComparison.Ordinal) < 0,
                    shared);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunGallery(TestRun run, WdSession session)
        {
            if (!Need(run, "tu-update-gallery", "tu-delete-wrap"))
            {
                return;
            }

            if (!Open(session, "后文你好啊。", out string path, out string channel) || !InsertGallery(session, path))
            {
                FailMany(run, "铺不出目录控件", "tu-update-gallery", "tu-delete-wrap");
                return;
            }

            try
            {
                string display = Paragraph(session, channel);
                Match mark = TocToken.Match(display);
                if (!mark.Success)
                {
                    Fail(run, "tu-update-gallery", "display 里没有目录编号：" + display);
                    Fail(run, "tu-delete-wrap", "display 里没有目录编号");
                    return;
                }

                string id = mark.Groups[1].Value;
                string updated = GetXml(session, channel, id, "<toc><content><levels>1-1</levels></content></toc>");
                string again = Paragraph(session, channel);
                Ok(run, "tu-update-gallery",
                    updated.Contains("<levels>1-1</levels>") && again.Contains(id),
                    updated);

                ToolResult deleted = Toc(channel, "delete", id, null, null);
                string after = Paragraph(session, channel);
                Ok(run, "tu-delete-wrap",
                    deleted.Success && after.IndexOf(id, StringComparison.Ordinal) < 0
                    && after.IndexOf("键入章标题", StringComparison.Ordinal) < 0
                    && after.Contains("后文你好啊"),
                    after);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunDeletePlain(TestRun run, WdSession session)
        {
            if (!Want("tu-delete-plain"))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel))
            {
                Fail(run, "tu-delete-plain", "打不开临时稿");
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null,
                    "<toc><format><title text=\"目录\"/></format></toc>", null);
                string id = WdSession.Str(inserted, "toc_id");
                ToolResult deleted = Toc(channel, "delete", id, null, null);
                string display = Paragraph(session, channel);
                Ok(run, "tu-delete-plain",
                    deleted.Success && display.IndexOf(id, StringComparison.Ordinal) < 0 && display.Contains("目录"),
                    display);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunEmpty(TestRun run, WdSession session)
        {
            if (!Want("tu-empty-ok"))
            {
                return;
            }

            if (!Open(session, "没有标题的正文。", out string path, out string channel))
            {
                Fail(run, "tu-empty-ok", "打不开临时稿");
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null, null, null);
                Ok(run, "tu-empty-ok",
                    inserted.Success && WdSession.Str(inserted, "toc_id").StartsWith("ewC_", StringComparison.Ordinal),
                    inserted.Success ? WdSession.Str(inserted, "message") : inserted.Error);
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void RunKeys(TestRun run, WdSession session)
        {
            if (!NeedKeys(run))
            {
                return;
            }

            if (!Open(session, "第一章标题\r正文停留。", out string path, out string channel) || !StyleHeading(session, path))
            {
                FailMany(run, "铺不出带标题的临时稿", KeyIds());
                return;
            }

            try
            {
                ToolResult inserted = Toc(channel, "insert", null, null, null);
                string id = WdSession.Str(inserted, "toc_id");
                if (!inserted.Success)
                {
                    FailMany(run, inserted.Error ?? "插入失败", KeyIds());
                    return;
                }

                foreach (KeyCase item in KeyCases())
                {
                    if (!Want(item.Id))
                    {
                        continue;
                    }

                    if (item.Reject)
                    {
                        ToolResult bad = Toc(channel, "update", id, item.Xml, null);
                        Ok(run, item.Id, !bad.Success && (bad.Error ?? "").Contains("不认识"), bad.Error ?? "");
                        continue;
                    }

                    if (item.Id == "tu-c-reflection")
                    {
                        ToolResult updated = Toc(channel, "update", id, item.Xml, null);
                        string warn = updated == null || updated.Data == null
                            ? ""
                            : Newtonsoft.Json.JsonConvert.SerializeObject(updated.Data);
                        ToolResult got = updated != null && updated.Success ? Toc(channel, "get", id, null, null) : null;
                        string readBack = got == null ? "" : (WdSession.Str(got, "xml") ?? "");
                        Ok(run, item.Id,
                            updated != null && updated.Success
                            && warn.IndexOf("预设倒影", StringComparison.Ordinal) >= 0
                            && readBack.IndexOf("size=\"20\"", StringComparison.Ordinal) >= 0
                            && readBack.IndexOf("msoReflectionType1", StringComparison.Ordinal) < 0,
                            warn + "\n" + readBack);
                        continue;
                    }

                    if (item.Id == "tu-c-no-proofing" && session.IsWps)
                    {
                        ToolResult updated = Toc(channel, "update", id, item.Xml, null);
                        Ok(run, item.Id,
                            !updated.Success && (updated.Error ?? "").IndexOf("no_proofing", StringComparison.OrdinalIgnoreCase) >= 0,
                            updated.Error ?? "");
                        continue;
                    }

                    string xml = GetXml(session, channel, id, item.Xml);
                    Ok(run, item.Id, xml.IndexOf(item.Expect, StringComparison.OrdinalIgnoreCase) >= 0, xml);
                }
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void OkXml(TestRun run, WdSession session, string channel, string id, string caseId, string xml, string expect)
        {
            if (!Want(caseId))
            {
                return;
            }

            string got = GetXml(session, channel, id, xml);
            Ok(run, caseId, got.IndexOf(expect, StringComparison.Ordinal) >= 0, got);
        }

        private static void OkLeader(TestRun run, WdSession session, string channel, string id, string caseId, string name, string enumName)
        {
            if (!Want(caseId))
            {
                return;
            }

            string got = GetXml(session, channel, id, "<toc><format><leader>" + name + "</leader></format></toc>");
            Ok(run, caseId,
                got.IndexOf(name, StringComparison.Ordinal) >= 0 || got.IndexOf(enumName, StringComparison.Ordinal) >= 0,
                got);
        }

        private static string GetXml(WdSession session, string channel, string id, string xml)
        {
            ToolResult updated = Toc(channel, "update", id, xml, null);
            if (updated == null)
            {
                return "没有回包";
            }

            if (!updated.Success)
            {
                return updated.Error ?? "更新失败";
            }

            ToolResult got = Toc(channel, "get", id, null, null);
            if (got == null || !got.Success)
            {
                return got == null ? "没有回包" : (got.Error ?? "读失败");
            }

            return WdSession.Str(got, "xml");
        }

        private static void CheckFail(
            TestRun run,
            WdSession session,
            string channel,
            string caseId,
            string action,
            string tocId,
            string xml,
            string at,
            string expect)
        {
            if (!Want(caseId))
            {
                return;
            }

            ToolResult result = Toc(channel, action, tocId, xml, at);
            string error = result == null ? "没有回包" : (result.Error ?? "");
            Ok(run, caseId, result != null && !result.Success && error.Contains(expect), error);
        }

        private static ToolResult Toc(string channel, string action, string tocId, string xml, string at)
        {
            var args = new Dictionary<string, object>();
            if (!string.IsNullOrEmpty(channel))
            {
                args["channel_id"] = channel;
            }

            if (action != null)
            {
                args["action"] = action;
            }

            if (!string.IsNullOrEmpty(tocId))
            {
                args["toc_id"] = tocId;
            }

            if (xml != null)
            {
                args["xml"] = xml;
            }

            if (!string.IsNullOrEmpty(at))
            {
                args["at"] = at;
            }

            return DesktopToolPipe.Call("F_toc", args);
        }

        private static bool Open(WdSession session, string plain, out string path, out string channel)
        {
            path = session.NewDocxPath();
            channel = null;
            ToolResult opened = session.Open(path, true, session.App);
            if (opened == null || !opened.Success)
            {
                return false;
            }

            channel = WdSession.Str(opened, "channel_id");
            return channel.Length > 0 && session.TypePlain(path, plain);
        }

        private static bool StyleHeading(WdSession session, string path)
        {
            string[] names = { "标题 1", "heading 1", "Heading 1" };
            if (session.IsWps)
            {
                return session.TryWpsDoc(path, doc =>
                {
                    object paragraph = FirstParagraph(doc);
                    object styles = WdSession.ComGet(doc, "Styles");
                    object style = null;
                    foreach (string name in names)
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

                    if (paragraph == null || style == null)
                    {
                        return false;
                    }

                    WdSession.ComSet(paragraph, "Style", style);
                    try
                    {
                        WdSession.ComSet(paragraph, "OutlineLevel", 1);
                    }
                    catch (Exception)
                    {
                    }

                    return true;
                });
            }

            return session.OnWordDoc(path, doc =>
            {
                if (doc.Paragraphs.Count < 1)
                {
                    return false;
                }

                foreach (string name in names)
                {
                    try
                    {
                        object style = doc.Styles[name];
                        doc.Paragraphs[1].set_Style(ref style);
                        try
                        {
                            doc.Paragraphs[1].OutlineLevel = Word.WdOutlineLevel.wdOutlineLevel1;
                        }
                        catch (Exception)
                        {
                        }

                        return true;
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
        }

        private static object FirstParagraph(object doc)
        {
            object paragraphs = WdSession.ComGet(doc, "Paragraphs");
            try
            {
                return WdSession.ComCall(paragraphs, "Item", 1);
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static bool InsertGallery(WdSession session, string path)
        {
            const string xml =
                "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:t>键入章标题(第 1 级)</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>后文你好啊。</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
            if (session.IsWps)
            {
                return session.TryWpsDoc(path, doc =>
                {
                    object content = WdSession.ComGet(doc, "Content");
                    WdSession.ComCall(content, "InsertXML", xml);
                    return true;
                });
            }

            return session.OnWordDoc(path, doc =>
            {
                doc.Content.InsertXML(xml);
                return true;
            });
        }

        private static bool StyleSizeIsNot(WdSession session, string path, float banned)
        {
            float size = ReadStyleSize(session, path);
            return size > 0 && Math.Abs(size - banned) > 0.2f;
        }

        private static bool SetTocStyleSize(WdSession session, string path, float size)
        {
            string[] names = { "目录 1", "TOC 1" };
            if (session.IsWps)
            {
                return session.TryWpsDoc(path, doc =>
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    foreach (string name in names)
                    {
                        try
                        {
                            object style = WdSession.ComCall(styles, "Item", name);
                            object font = WdSession.ComGet(style, "Font");
                            WdSession.ComSet(font, "Size", size);
                            return true;
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return false;
                });
            }

            return session.OnWordDoc(path, doc =>
            {
                foreach (string name in names)
                {
                    try
                    {
                        doc.Styles[name].Font.Size = size;
                        return true;
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
        }

        private static float ReadStyleSize(WdSession session, string path)
        {
            float found = 0;
            string[] names = { "目录 1", "TOC 1" };
            if (session.IsWps)
            {
                session.TryWpsDoc(path, doc =>
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    foreach (string name in names)
                    {
                        try
                        {
                            object style = WdSession.ComCall(styles, "Item", name);
                            object font = WdSession.ComGet(style, "Font");
                            object size = WdSession.ComGet(font, "Size");
                            found = Convert.ToSingle(size);
                            return found > 0;
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return false;
                });
                return found;
            }

            session.OnWordDoc(path, doc =>
            {
                foreach (string name in names)
                {
                    try
                    {
                        found = doc.Styles[name].Font.Size;
                        return true;
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
            return found;
        }

        private static string Paragraph(WdSession session, string channel)
        {
            ToolResult got = session.GetContent(channel, "paragraph");
            return got != null && got.Success ? WdSession.Str(got, "document_content") : (got == null ? "" : got.Error ?? "");
        }

        private static List<string> ParaCodes(WdSession session, string channel)
        {
            var codes = new List<string>();
            foreach (Match match in ParaCode.Matches(Paragraph(session, channel)))
            {
                if (!codes.Contains(match.Value))
                {
                    codes.Add(match.Value);
                }
            }

            return codes;
        }

        private static string FirstInteriorParagraph(WdSession session, string path)
        {
            string found = "";
            if (session.IsWps)
            {
                session.TryWpsDoc(path, doc =>
                {
                    Word.Document wordDoc = doc as Word.Document;
                    if (wordDoc == null)
                    {
                        return false;
                    }

                    WordDocumentExtractor.ProcessDocument(wordDoc, ProcessDocumentOptions.ForProcessActions("locate"));
                    found = TocFieldStamp.FirstInteriorParagraph() ?? "";
                    return true;
                });
            }
            else
            {
                session.OnWordDoc(path, doc =>
                {
                    WordDocumentExtractor.ProcessDocument(doc, ProcessDocumentOptions.ForProcessActions("locate"));
                    found = TocFieldStamp.FirstInteriorParagraph() ?? "";
                    return true;
                });
            }

            return found;
        }

        private static string LineText(ToolResult result)
        {
            Dictionary<string, object> data = WdSession.Data(result);
            if (data == null || !data.TryGetValue("lines", out object raw) || raw == null)
            {
                return "";
            }

            var buffer = new List<string>();
            if (raw is JArray array)
            {
                foreach (JToken row in array)
                {
                    buffer.Add(row["text"]?.ToString() ?? "");
                }
            }
            else if (raw is IEnumerable list && !(raw is string))
            {
                foreach (object row in list)
                {
                    if (row is JObject obj)
                    {
                        buffer.Add(obj["text"]?.ToString() ?? "");
                    }
                    else if (row is IDictionary map && map.Contains("text"))
                    {
                        buffer.Add(Convert.ToString(map["text"]) ?? "");
                    }
                }
            }

            return string.Join("\n", buffer);
        }

        private static IEnumerable<string> AllIds()
        {
            yield return "tu-x-action";
            yield return "tu-x-no-action";
            yield return "tu-x-no-id";
            yield return "tu-x-bad-id";
            yield return "tu-x-no-xml";
            yield return "tu-x-root";
            yield return "tu-x-levels";
            yield return "tu-x-style";
            yield return "tu-x-sep";
            yield return "tu-x-leader";
            yield return "tu-x-font";
            yield return "tu-x-range";
            yield return "tu-levels-default";
            yield return "tu-levels-11";
            yield return "tu-levels-19";
            yield return "tu-levels-99";
            yield return "tu-styles-one";
            yield return "tu-styles-clear";
            yield return "tu-styles-keep";
            yield return "tu-range-pair";
            yield return "tu-range-clear";
            yield return "tu-range-half";
            yield return "tu-page-off";
            yield return "tu-page-49";
            yield return "tu-align-after";
            yield return "tu-leader-none";
            yield return "tu-leader-dash";
            yield return "tu-leader-line";
            yield return "tu-leader-heavy";
            yield return "tu-leader-middle";
            yield return "tu-hyper-off";
            yield return "tu-web-off";
            yield return "tu-chapter";
            yield return "tu-keep-tabs";
            yield return "tu-keep-breaks";
            yield return "tu-sep-five";
            yield return "tu-title-on";
            yield return "tu-title-off";
            yield return "tu-title-level";
            yield return "tu-get-one";
            yield return "tu-get-text";
            yield return "tu-get-omit";
            yield return "tu-insert-at";
            yield return "tu-insert-interior";
            yield return "tu-update-part";
            yield return "tu-format-local";
            yield return "tu-update-gallery";
            yield return "tu-delete-wrap";
            yield return "tu-delete-plain";
            yield return "tu-reset";
            yield return "tu-style-shared";
            yield return "tu-empty-ok";
            foreach (KeyCase item in KeyCases())
            {
                yield return item.Id;
            }
        }

        private static string[] KeyIds()
        {
            var ids = new List<string>();
            foreach (KeyCase item in KeyCases())
            {
                ids.Add(item.Id);
            }

            return ids.ToArray();
        }

        private static bool NeedKeys(TestRun run)
        {
            foreach (KeyCase item in KeyCases())
            {
                if (Want(item.Id))
                {
                    return true;
                }
            }

            return false;
        }

        private static IEnumerable<KeyCase> KeyCases()
        {
            foreach (string[] row in CharRows())
            {
                yield return Key("tu-c-" + row[0].Replace('_', '-'), "font", row[0], row[1], row[2]);
            }

            foreach (string[] row in ParaRows())
            {
                yield return Key("tu-p-" + row[0].Replace('_', '-'), "paragraph", row[0], row[1], row[2]);
            }

            yield return new KeyCase
            {
                Id = "tu-p-reject-style",
                Xml = "<toc><format><level n=\"1\"><paragraph paragraph_style=\"标题 1\"/></level></format></toc>",
                Expect = "不认识",
                Reject = true
            };
            for (int level = 2; level <= 9; level++)
            {
                yield return new KeyCase
                {
                    Id = "tu-lv-" + level,
                    Xml = "<toc><format><level n=\"" + level + "\"><font font_size=\"18\"/></level></format></toc>",
                    Expect = "n=\"" + level + "\""
                };
            }

            yield return new KeyCase
            {
                Id = "tu-lv-title",
                Xml = "<toc><format><title text=\"目录\"/><level n=\"title\"><font font_name=\"黑体\"/></level></format></toc>",
                Expect = "n=\"title\""
            };
        }

        private static KeyCase Key(string id, string side, string name, string body, string expect)
        {
            string inner = body.Length == 0 ? name + "=\"" + expect + "\"" : body;
            string xml = body.Length == 0
                ? "<toc><format><level n=\"1\"><" + side + " " + inner + "/></level></format></toc>"
                : "<toc><format><level n=\"1\"><" + side + ">" + inner + "</" + side + "></level></format></toc>";
            return new KeyCase { Id = id, Xml = xml, Expect = expect };
        }

        private static string[][] CharRows()
        {
            return new[]
            {
                new[] { "font_name", "", "黑体" },
                new[] { "name", "", "宋体" },
                new[] { "name_ascii", "", "Arial" },
                new[] { "name_far_east", "", "宋体" },
                new[] { "name_bi", "", "Arial" },
                new[] { "name_other", "", "Arial" },
                new[] { "font_size", "", "16" },
                new[] { "size_bi", "", "16" },
                new[] { "font_color", "", "#FF0000" },
                new[] { "color_index", "", "wdRed" },
                new[] { "color_index_bi", "", "wdBlue" },
                new[] { "text_color", "<text_color rgb=\"#FF0000\"/>", "FF0000" },
                new[] { "underline_color", "", "#0000FF" },
                new[] { "diacritic_color", "", "#00FF00" },
                new[] { "is_bold", "", "true" },
                new[] { "bold_bi", "", "true" },
                new[] { "is_italic", "", "true" },
                new[] { "italic_bi", "", "true" },
                new[] { "has_underline", "", "true" },
                new[] { "underline", "", "wdUnderlineSingle" },
                new[] { "has_strikethrough", "", "true" },
                new[] { "double_strikethrough", "", "true" },
                new[] { "is_subscript", "", "true" },
                new[] { "is_superscript", "", "true" },
                new[] { "shadow", "", "true" },
                new[] { "outline", "", "true" },
                new[] { "emboss", "", "true" },
                new[] { "engrave", "", "true" },
                new[] { "small_caps", "", "true" },
                new[] { "all_caps", "", "true" },
                new[] { "emphasis_mark", "", "wdEmphasisMarkOverSolidCircle" },
                new[] { "scaling", "", "150" },
                new[] { "spacing", "", "1" },
                new[] { "position", "", "2" },
                new[] { "kerning", "", "12" },
                new[] { "disable_character_space_grid", "", "true" },
                new[] { "ligatures", "", "wdLigaturesStandard" },
                new[] { "number_form", "", "wdNumberFormLining" },
                new[] { "number_spacing", "", "wdNumberSpacingProportional" },
                new[] { "stylistic_set", "", "wdStylisticSet01" },
                new[] { "contextual_alternates", "", "true" },
                new[] { "background_color", "", "#FFFF00" },
                new[] { "shading", "<shading texture=\"wdTexture5Percent\" background_pattern_color=\"#00FF00\"/>", "wdTexture5Percent" },
                new[] { "highlight_color_index", "", "wdYellow" },
                new[] { "borders", "<borders outside_line_style=\"wdLineStyleSingle\" outside_color=\"#FF0000\"/>", "wdLineStyleSingle" },
                new[] { "fill", "<fill visible=\"true\" transparency=\"0\"/>", "visible" },
                new[] { "glow", "<glow radius=\"2\" transparency=\"0\"/>", "radius" },
                new[] { "line", "<line visible=\"true\" weight=\"1\"/>", "weight" },
                new[] { "reflection", "<reflection type=\"msoReflectionType1\" size=\"20\"/>", "msoReflectionType1" },
                new[] { "text_shadow", "<text_shadow visible=\"true\" offset_x=\"2\" offset_y=\"2\"/>", "offset_x" },
                new[] { "language_id", "", "wdSimplifiedChinese" },
                new[] { "language_id_far_east", "", "wdSimplifiedChinese" },
                new[] { "no_proofing", "", "true" },
                new[] { "fit_text_width", "", "100" },
            };
        }

        private static string[][] ParaRows()
        {
            return new[]
            {
                new[] { "alignment", "", "center" },
                new[] { "left_indent", "", "12" },
                new[] { "right_indent", "", "12" },
                new[] { "first_line_indent", "", "21" },
                new[] { "hanging_indent", "", "21" },
                new[] { "space_before", "", "6" },
                new[] { "space_after", "", "6" },
                new[] { "line_spacing_rule", "", "wdLineSpace1pt5" },
                new[] { "line_spacing", "", "18" },
                new[] { "outline_level", "", "wdOutlineLevel2" },
                new[] { "baseline_alignment", "", "wdBaselineAlignCenter" },
                new[] { "reading_order", "", "wdReadingOrderLtr" },
                new[] { "mirror_indents", "", "true" },
                new[] { "auto_adjust_right_indent", "", "false" },
                new[] { "character_unit_left_indent", "", "1" },
                new[] { "character_unit_right_indent", "", "1" },
                new[] { "character_unit_first_line_indent", "", "1" },
                new[] { "line_unit_before", "", "1" },
                new[] { "line_unit_after", "", "1" },
                new[] { "space_before_auto", "", "true" },
                new[] { "space_after_auto", "", "true" },
                new[] { "disable_line_height_grid", "", "true" },
                new[] { "widow_control", "", "false" },
                new[] { "keep_with_next", "", "true" },
                new[] { "keep_together", "", "true" },
                new[] { "page_break_before", "", "true" },
                new[] { "no_line_number", "", "true" },
                new[] { "hyphenation", "", "false" },
                new[] { "far_east_line_break_control", "", "false" },
                new[] { "word_wrap", "", "false" },
                new[] { "hanging_punctuation", "", "true" },
                new[] { "half_width_punctuation_on_top_of_line", "", "true" },
                new[] { "add_space_between_far_east_and_alpha", "", "false" },
                new[] { "add_space_between_far_east_and_digit", "", "false" },
                new[] { "textbox_tight_wrap", "", "wdTightAll" },
                new[] { "collapsed_by_default", "", "true" },
                new[] { "shading", "<shading texture=\"wdTexture5Percent\" background_pattern_color=\"#00FF00\"/>", "00FF00" },
                new[] { "borders", "<borders><top line_style=\"wdLineStyleSingle\" line_width=\"wdLineWidth050pt\"/></borders>", "wdLineStyleSingle" },
                new[] { "tab_stops", "<tab_stops><tab position=\"72\" alignment=\"left\" leader=\"none\"/></tab_stops>", "72" },
                new[] { "list", "<list kind=\"bullet\"/>", "bullet" },
                new[] { "drop_cap", "<drop_cap position=\"drop\" lines=\"3\"/>", "drop" },
            };
        }

        private static bool Need(TestRun run, params string[] ids)
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
            return Only == null || Only.Contains(id);
        }

        private static void Ok(TestRun run, string id, bool pass, string detail)
        {
            if (!Want(id))
            {
                return;
            }

            Seen.Add(id);
            if (pass)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, detail ?? "");
            }
        }

        private static void Fail(TestRun run, string id, string reason)
        {
            Ok(run, id, false, reason);
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                Fail(run, id, reason);
            }
        }

        private sealed class KeyCase
        {
            public string Id;
            public string Xml;
            public string Expect;
            public bool Reject;
        }
    }
}
