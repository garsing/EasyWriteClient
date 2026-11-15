using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>样式工具。对照 docs/测试/word测试/样式工具。铺稿用 COM，列出和套用走 F_document_structure。</summary>
    internal static class DsCases
    {
        public const int Target = 91;

        private static readonly string[] HeadingNames = { "标题 1", "heading 1", "Heading 1" };
        private static readonly string[] HeadingCharNames = { "标题 1 字符", "heading 1 Char", "Heading 1 Char" };
        private static readonly string[] BodyNames = { "正文", "Normal" };
        private static readonly string[] ListNames = { "列表段落", "List Paragraph" };
        private static readonly string[] LinkNames = { "超链接", "Hyperlink" };
        private static readonly string[] Heading2Names = { "标题 2", "heading 2", "Heading 2" };
        private static readonly string[] Heading9Names = { "标题 9", "heading 9", "Heading 9" };
        private static readonly string[] TitleNames = { "标题", "Title" };
        private static readonly string[] SubtitleNames = { "副标题", "Subtitle" };

        private static HashSet<string> Only;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = null;
            if (!string.IsNullOrWhiteSpace(filter)
                && !filter.Equals("ds", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("ds-", StringComparison.OrdinalIgnoreCase))
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
            RunCatalog(run, s);
            RunSpans(run, s);
            RunApply(run, s);
            RunEdit(run, s);
            RunMore(run, s);
            RunToc(run, s);
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

        private static void RunCatalog(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "ds-span-need-display",
                "ds-cat-heading-pair", "ds-cat-heading-format", "ds-cat-body-defaults", "ds-cat-body-empty-rpr",
                "ds-cat-list-para", "ds-cat-hyperlink", "ds-cat-title", "ds-cat-subtitle",
                "ds-cat-heading-unused", "ds-cat-quote", "ds-cat-emphasis", "ds-cat-def-font",
                "ds-cat-toc", "ds-cat-caption-header", "ds-cat-rules", "ds-cat-display-intact"
            };
            if (!Want(ids))
            {
                return;
            }
            string path, id;
            if (!Open(s, "样式目录甲。", out path, out id))
            {
                FailMany(run, "铺目录稿失败", ids);
                return;
            }

            try
            {
                JObject early = Act(s, id, "list_spans");
                bool needFail = !OkFlag(early) && Blob(early, "F_get_document_content");
                Ok(run, "ds-span-need-display", needFail, Dump(early));
                Materialize(s, path);

                JObject listed = Act(s, id, "list_styles");
                JArray styles = Arr(listed, "styles");
                JObject heading = Find(styles, "heading_1");
                JObject body = Find(styles, "body");
                bool headingPair = heading != null && Eq(heading, "origin", "builtin")
                    && Find(styles, "heading 1 Char") == null && Find(styles, "标题 1 字符") == null
                    && Eq(heading, "support", Obj(heading, "text_format") != null ? "both" : "paragraph");
                Ok(run, "ds-cat-heading-pair", headingPair, Dump(heading));
                JObject hText = Obj(heading, "text_format");
                JObject hPara = Obj(heading, "paragraph_format");
                bool headingFormat = s.IsWps
                    ? heading != null && hText != null && Near(hText, "font_size", 22, 1.2) && Flag(hText, "is_bold")
                        && Eq(hPara, "outline_level", "wdOutlineLevel1")
                        && hPara["space_before"] != null && hPara["space_after"] != null
                    : heading != null && hText != null && hText["font_size"] != null && hText["font_color"] != null
                        && hText["is_bold"] == null
                        && Eq(hPara, "outline_level", "wdOutlineLevel1")
                        && (hPara["space_before"] != null || hPara["space_after"] != null);
                Ok(run, "ds-cat-heading-format", headingFormat, Dump(heading));
                JObject bodyText = Obj(body, "text_format");
                JObject bodyPara = Obj(body, "paragraph_format");
                bool bodyDefaults = s.IsWps
                    ? body != null && Eq(body, "support", "both")
                        && Near(bodyText, "font_size", 10.5, 0.6)
                        && bodyText["is_bold"] == null
                        && (bodyPara == null || bodyPara["outline_level"] == null)
                    : body != null && Eq(body, "support", "both")
                        && Near(bodyText, "font_size", 11, 0.6)
                        && Near(bodyPara, "space_after", 8, 1.2)
                        && bodyText["is_bold"] == null
                        && bodyPara["outline_level"] == null;
                Ok(run, "ds-cat-body-defaults", bodyDefaults, Dump(body));
                Ok(run, "ds-cat-body-empty-rpr",
                    body != null && Obj(body, "text_format") != null && Obj(body, "text_format")["font_size"] != null,
                    Dump(body));
                JObject list = Find(styles, "列表段落");
                bool listBuiltin = HostHasStyle(s, path, ListNames);
                Ok(run, "ds-cat-list-para",
                    listBuiltin
                        ? list != null && Eq(list, "support", "paragraph") && Near(Obj(list, "paragraph_format"), "left_indent", 36, 1.2) && Obj(list, "text_format") == null
                        : list == null,
                    Dump(list));
                JObject link = Find(styles, "超链接");
                string color = Str(Obj(link, "text_format"), "font_color");
                string linkColor = s.IsWps ? "0000FF" : "467886";
                Ok(run, "ds-cat-hyperlink",
                    link != null && Eq(link, "support", "character") && color.IndexOf(linkColor, StringComparison.OrdinalIgnoreCase) >= 0
                    && Flag(Obj(link, "text_format"), "has_underline") && Obj(link, "paragraph_format") == null,
                    Dump(link));
                JObject title = Find(styles, "title");
                Ok(run, "ds-cat-title",
                    title != null && Eq(title, "support", Obj(title, "text_format") != null ? "both" : "paragraph"),
                    Dump(title));
                JObject subtitle = Find(styles, "subtitle");
                string subOutline = Str(Obj(subtitle, "paragraph_format"), "outline_level");
                bool subtitleOk = subtitle != null && !Eq(subtitle, "structure", "heading_2")
                    && (s.IsWps
                        ? Eq(subtitle, "structure", "subtitle") && subOutline.IndexOf("2", StringComparison.Ordinal) >= 0
                        : subOutline.Length == 0 || subOutline.IndexOf("2", StringComparison.Ordinal) < 0);
                Ok(run, "ds-cat-subtitle", subtitleOk, Dump(subtitle));
                bool unused = true;
                for (int n = 2; n <= 9; n++)
                {
                    if (Find(styles, "heading_" + n) == null)
                    {
                        unused = false;
                    }
                }

                Ok(run, "ds-cat-heading-unused", unused, "缺标题 2～9");
                JObject quote = Find(styles, "引用");
                JObject intense = Find(styles, "明显引用");
                bool quoteBuiltin = HostHasStyle(s, path, "引用", "Quote") && HostHasStyle(s, path, "明显引用", "Intense Quote");
                Ok(run, "ds-cat-quote",
                    quoteBuiltin
                        ? quote != null && Eq(quote, "support", "both") && intense != null && Eq(intense, "support", "both")
                        : quote == null && intense == null,
                    Dump(quote) + " / " + Dump(intense));
                JObject emphasis = Find(styles, "明显强调");
                bool emphasisBuiltin = HostHasStyle(s, path, "明显强调", "Intense Emphasis");
                Ok(run, "ds-cat-emphasis",
                    emphasisBuiltin
                        ? emphasis != null && Eq(emphasis, "support", "character") && Eq(emphasis, "origin", "builtin")
                        : emphasis == null,
                    Dump(emphasis));
                JObject defFont = Find(styles, "默认段落字体");
                Ok(run, "ds-cat-def-font", defFont != null && Eq(defFont, "support", "character"), Dump(defFont));
                Ok(run, "ds-cat-toc", TocOk(styles), "没有目录样式");
                JObject caption = Find(styles, "题注") ?? Find(styles, "Caption");
                JObject header = Find(styles, "页眉") ?? Find(styles, "Header");
                JObject footer = Find(styles, "页脚") ?? Find(styles, "Footer");
                Ok(run, "ds-cat-caption-header",
                    caption != null && Eq(caption, "support", Obj(caption, "text_format") != null ? "both" : "paragraph")
                    && header != null && Eq(header, "support", Obj(header, "text_format") != null ? "both" : "paragraph")
                    && footer != null && Eq(footer, "support", Obj(footer, "text_format") != null ? "both" : "paragraph"),
                    "题注=" + Dump(caption) + " 页眉=" + Dump(header) + " 页脚=" + Dump(footer));
                JObject rules = Obj(listed, "apply_rules");
                Ok(run, "ds-cat-rules",
                    rules != null && Str(rules, "both").IndexOf("段落格式没有发生作用", StringComparison.Ordinal) >= 0
                    && Str(rules, "paragraph").IndexOf("什么也不会发生", StringComparison.Ordinal) >= 0
                    && rules["character"] != null && ApplyMatches(styles),
                    Dump(rules));

                string before = Content(s, id);
                Act(s, id, "list_styles");
                string after = Content(s, id);
                Ok(run, "ds-cat-display-intact",
                    before.Length > 0 && before == after && before.IndexOf("structure=", StringComparison.Ordinal) < 0 && before.IndexOf("[st ", StringComparison.Ordinal) < 0,
                    Trim(before) + " -> " + Trim(after));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpans(TestRun run, WdSession s)
        {
            if (Want("ds-span-heading-text", "ds-span-body-rest", "ds-span-two")) RunSpanPair(run, s);
            if (Want("ds-span-whole-omit")) RunSpanWhole(run, s);
            if (Want("ds-span-direct-color")) RunSpanColor(run, s);
            if (Want("ds-span-merge")) RunSpanMerge(run, s);
            if (Want("ds-span-repeat")) RunSpanRepeat(run, s);
            if (Want("ds-span-list-scope")) RunSpanList(run, s);
            if (Want("ds-span-link-scope")) RunSpanLink(run, s);
            if (Want("ds-span-toc-skip")) RunSpanToc(run, s);
            if (Want("ds-span-image")) RunSpanImage(run, s);
            if (Want("ds-span-miss")) RunSpanMiss(run, s);
            if (Want("ds-span-empty")) RunSpanEmpty(run, s);
            if (Want("ds-span-custom-name")) RunSpanCustom(run, s);
            if (Want("ds-span-table")) RunSpanTable(run, s);
        }

        private static void RunSpanPair(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "你好啊哈哈哈哈", out path, out id))
            {
                FailMany(run, "铺分截稿失败", "ds-span-heading-text", "ds-span-body-rest", "ds-span-two");
                return;
            }

            try
            {
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                bool painted = headingChar && PaintSlice(s, path, "你好啊哈哈哈哈", "你好啊", HeadingCharNames);
                string body = Content(s, id);
                JObject listed = Act(s, id, "list_spans");
                JArray spans = Arr(listed, "spans");
                JObject head = SpanText(spans, "你好啊");
                JObject rest = SpanText(spans, "哈哈哈哈");
                Ok(run, "ds-span-heading-text",
                    !headingChar || (painted && head != null && Eq(head, "structure", "heading_1") && Eq(head, "scope", "text")
                    && Obj(head, "text_format") != null && Obj(head, "paragraph_format") == null),
                    painted ? Dump(head) : "铺标题 1 字符失败");
                Ok(run, "ds-span-body-rest",
                    !headingChar || (rest != null && Eq(rest, "structure", "body") && Eq(rest, "scope", "paragraph")
                    && Obj(rest, "text_format") != null && Obj(rest, "paragraph_format") != null),
                    Dump(rest));
                Ok(run, "ds-span-two",
                    !headingChar || (head != null && rest != null && Str(head, "target_codes") == Str(rest, "target_codes")
                    && Str(head, "target_codes").StartsWith("S_", StringComparison.Ordinal)
                    && body.IndexOf(Str(head, "target_codes"), StringComparison.Ordinal) >= 0),
                    Dump(head) + " / " + Dump(rest));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanWhole(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "整句只有正文。", out path, out id))
            {
                FailMany(run, "铺整句稿失败", "ds-span-whole-omit");
                return;
            }

            try
            {
                Content(s, id);
                JObject one = First(Arr(Act(s, id, "list_spans"), "spans"));
                Ok(run, "ds-span-whole-omit",
                    one != null && Eq(one, "structure", "body") && Eq(one, "scope", "paragraph") && one["text"] == null,
                    Dump(one));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanColor(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "涂红一句。", out path, out id))
            {
                FailMany(run, "铺涂红稿失败", "ds-span-direct-color");
                return;
            }

            try
            {
                bool painted = PaintColor(s, path, "涂红一句。", "红");
                Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                Ok(run, "ds-span-direct-color",
                    painted && spans.Count == 1 && Dump(First(spans)).IndexOf("FF0000", StringComparison.OrdinalIgnoreCase) < 0,
                    painted ? Dump(First(spans)) : "涂红失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanMerge(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "甲乙丙丁", out path, out id))
            {
                FailMany(run, "铺合并稿失败", "ds-span-merge");
                return;
            }

            try
            {
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                bool a = headingChar && PaintSlice(s, path, "甲乙丙丁", "甲乙", HeadingCharNames);
                bool b = headingChar && PaintSlice(s, path, "甲乙丙丁", "丙丁", HeadingCharNames);
                Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                JObject merged = SpanText(spans, "甲乙丙丁");
                Ok(run, "ds-span-merge",
                    !headingChar || (a && b && merged != null && Eq(merged, "structure", "heading_1") && spans.Count == 1),
                    Dump(spans));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanRepeat(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "好好好", out path, out id))
            {
                FailMany(run, "铺重复稿失败", "ds-span-repeat");
                return;
            }

            try
            {
                bool painted = PaintNth(s, path, "好好好", "好", 2, LinkNames);
                Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                JObject hit = null;
                foreach (JObject span in Each(spans))
                {
                    if (Eq(span, "structure", "超链接") && TextIndex(span) == 2 && TextValue(span) == "好")
                    {
                        hit = span;
                    }
                }

                Ok(run, "ds-span-repeat", painted && hit != null, painted ? Dump(spans) : "第二处超链接没铺上");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanList(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "列表整段。", out path, out id))
            {
                FailMany(run, "铺列表稿失败", "ds-span-list-scope");
                return;
            }

            try
            {
                bool listBuiltin = HostHasStyle(s, path, ListNames);
                bool painted = listBuiltin && PaintParagraph(s, path, "列表整段。", ListNames);
                Content(s, id);
                JObject one = First(Arr(Act(s, id, "list_spans"), "spans"));
                Ok(run, "ds-span-list-scope",
                    listBuiltin
                        ? painted && one != null && Eq(one, "scope", "paragraph") && Obj(one, "text_format") == null
                            && Near(Obj(one, "paragraph_format"), "left_indent", 36, 1.2)
                        : !painted,
                    painted ? Dump(one) : "铺列表段落失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanLink(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "链接几个字。", out path, out id))
            {
                FailMany(run, "铺链接稿失败", "ds-span-link-scope");
                return;
            }

            try
            {
                bool painted = PaintSlice(s, path, "链接几个字。", "几个", LinkNames);
                Content(s, id);
                JObject one = SpanText(Arr(Act(s, id, "list_spans"), "spans"), "几个");
                Ok(run, "ds-span-link-scope",
                    painted && one != null && Eq(one, "structure", "超链接") && Eq(one, "scope", "text")
                    && Obj(one, "text_format") != null && Obj(one, "paragraph_format") == null,
                    painted ? Dump(one) : "铺超链接失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunToc(TestRun run, WdSession s)
        {
            if (Only == null)
            {
                return;
            }

            if (!Want(
                "ds-toc-token",
                "ds-toc-reuse",
                "ds-toc-notoc",
                "ds-toc-spans",
                "ds-toc-reject",
                "ds-toc-style"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r你好啊。", out path, out id))
            {
                FailMany(
                    run,
                    "铺目录稿失败",
                    "ds-toc-token",
                    "ds-toc-reuse",
                    "ds-toc-notoc",
                    "ds-toc-spans",
                    "ds-toc-reject",
                    "ds-toc-style");
                return;
            }

            try
            {
                bool painted = PaintParagraph(s, path, "第一章标题", HeadingNames);
                bool field = AddToc(s, path);
                bool updated = UpdateFields(s, path);
                bool planted = PlantTocBookmark(s, path);
                string body = Content(s, id);
                Match token = Regex.Match(body ?? "", @"\[目录域 (ewC_[A-Za-z0-9]{8,32})\]");
                string name = token.Success ? token.Groups[1].Value : "";
                bool bodyKept = (body ?? "").IndexOf("你好啊", StringComparison.Ordinal) >= 0;
                bool noEdge = (body ?? "").IndexOf(TocFieldStamp.Edge) < 0;
                bool glued = SentenceHolds(body, "你好啊", "\t");
                Ok(run, "ds-toc-token",
                    painted && field && updated && token.Success && bodyKept && noEdge && !glued,
                    "painted=" + painted + " field=" + field + " updated=" + updated + " body=" + body);

                string again = Content(s, id);
                Match token2 = Regex.Match(again ?? "", @"\[目录域 (ewC_[A-Za-z0-9]{8,32})\]");
                int marks = CountBookmarks(s, path, name);
                Ok(run, "ds-toc-reuse",
                    token.Success && token2.Success && token2.Groups[1].Value == name && marks == 1,
                    "first=" + name + " second=" + (token2.Success ? token2.Groups[1].Value : "") + " marks=" + marks);

                bool stillThere = BookmarkExists(s, path, "_Toc245561084");
                Ok(run, "ds-toc-notoc",
                    planted && stillThere && token.Success
                    && name.StartsWith("ewC_", StringComparison.Ordinal)
                    && !string.Equals(name, "_Toc245561084", StringComparison.Ordinal),
                    "name=" + name + " still=" + stillThere + " plant=" + PlantNote);

                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                string bodyCode = Sent(body, "你好啊");
                bool leaked = SpanHasTab(spans);
                Ok(run, "ds-toc-spans",
                    bodyCode.StartsWith("S_", StringComparison.Ordinal) && SpanByCode(spans, bodyCode) != null && !leaked,
                    "code=" + bodyCode + " leaked=" + leaked + " " + Dump(spans));

                string paras = ParagraphContent(s, id);
                string paraCode = TocParagraph(paras);
                bool rejected = false;
                string rejectDetail = "para=" + paraCode;
                if (paraCode.Length > 0)
                {
                    ToolResult bad = s.ApplyParagraphFormat(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["target_paragraph_codes"] = paraCode,
                        ["para_format"] = new Dictionary<string, object> { ["alignment"] = "center" }
                    });
                    string err = bad == null ? "" : (bad.Error ?? "");
                    rejected = bad != null && !bad.Success && err.IndexOf("目录域里面的编码", StringComparison.Ordinal) >= 0;
                    rejectDetail = err;
                }

                if (!s.IsWps)
                {
                    string sentenceErr = RejectInteriorSentence(s, path);
                    rejected = rejected || sentenceErr.IndexOf("目录域里面的编码", StringComparison.Ordinal) >= 0;
                    rejectDetail = rejectDetail + " | " + sentenceErr;
                }

                Ok(run, "ds-toc-reject", rejected, rejectDetail);

                Materialize(s, path);
                JArray styles = Arr(Act(s, id, "list_styles"), "styles");
                Ok(run, "ds-toc-style",
                    Find(styles, "目录 1") != null || Find(styles, "TOC 1") != null,
                    Dump(styles));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanToc(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "目录后正文。", out path, out id))
            {
                FailMany(run, "铺目录稿失败", "ds-span-toc-skip");
                return;
            }

            try
            {
                bool field = AddToc(s, path);
                string body = Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                string code = Sent(body, "目录后正文");
                Ok(run, "ds-span-toc-skip",
                    field && code.StartsWith("S_", StringComparison.Ordinal) && SpanByCode(spans, code) != null
                    && body.IndexOf("目录后正文", StringComparison.Ordinal) >= 0,
                    field ? Dump(spans) : "目录域没插上");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanImage(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "图前字。", out path, out id))
            {
                FailMany(run, "铺图稿失败", "ds-span-image");
                return;
            }

            try
            {
                bool picture = s.InsertPicture(path);
                bool typed = s.AppendText(path, "\r图后字。");
                string body = Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                string code = Sent(body, "图后字");
                Ok(run, "ds-span-image",
                    picture && typed && code.StartsWith("S_", StringComparison.Ordinal) && SpanByCode(spans, code) != null
                    && body.IndexOf("图后字", StringComparison.Ordinal) >= 0,
                    picture ? Dump(spans) : "插图失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanMiss(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "还在的句子。\r将对不上", out path, out id))
            {
                FailMany(run, "铺对不上稿失败", "ds-span-miss");
                return;
            }

            try
            {
                string body = Content(s, id);
                string keepCode = Sent(body, "还在的句子");
                bool changed = ReplaceText(s, path, "将对不上", "已经改掉");
                JObject listed = Act(s, id, "list_spans");
                string blob = Dump(listed);
                bool invented = false;
                foreach (Match m in Regex.Matches(blob, @"S_[0-9A-Za-z]{5}"))
                {
                    if (body.IndexOf(m.Value, StringComparison.Ordinal) < 0)
                    {
                        invented = true;
                    }
                }

                Ok(run, "ds-span-miss",
                    changed && Blob(listed, "对不上") && !invented && SpanByCode(Arr(listed, "spans"), keepCode) != null,
                    changed ? blob : "改字失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanEmpty(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "有字。\r", out path, out id))
            {
                FailMany(run, "铺空段稿失败", "ds-span-empty");
                return;
            }

            try
            {
                string body = Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                bool empty = false;
                foreach (JObject span in Each(spans))
                {
                    string text = TextValue(span);
                    if (text != null && text.Trim().Length == 0)
                    {
                        empty = true;
                    }
                }

                string code = Sent(body, "有字");
                Ok(run, "ds-span-empty", spans.Count >= 1 && !empty && SpanByCode(spans, code) != null, Dump(spans));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanCustom(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "自定义几个字。", out path, out id))
            {
                FailMany(run, "铺自定义稿失败", "ds-span-custom-name");
                return;
            }

            try
            {
                JObject created = Act(s, id, Map(
                    "action", "customize",
                    "create", true,
                    "name", "强调测列",
                    "text_format", Map("font_size", 16),
                    "paragraph_format", Map("outline_level", "wdOutlineLevel2")));
                bool applied = OkFlag(created) && PaintSlice(s, path, "自定义几个字。", "几个", new[] { "强调测列 字符", "强调测列" });
                Content(s, id);
                JObject hit = SpanText(Arr(Act(s, id, "list_spans"), "spans"), "几个");
                JObject row = Find(Arr(Act(s, id, "list_styles"), "styles"), "强调测列");
                string outline = Str(Obj(row, "paragraph_format"), "outline_level");
                Ok(run, "ds-span-custom-name",
                    applied && hit != null && Eq(hit, "structure", "强调测列")
                    && row != null && Eq(row, "structure", "强调测列")
                    && outline.IndexOf("2", StringComparison.Ordinal) >= 0,
                    applied ? Dump(hit) + Dump(row) : Dump(created));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunSpanTable(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "表外一句。", out path, out id))
            {
                FailMany(run, "铺表格稿失败", "ds-span-table");
                return;
            }

            try
            {
                bool table = s.InsertTable(path, 1, 1, new[] { "表内标题" });
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                bool painted = headingChar && PaintSlice(s, path, "表内标题", "表内", HeadingCharNames);
                s.OnWordDoc(path, doc =>
                {
                    doc.Save();
                    return true;
                });
                string body = Content(s, id);
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                string code = Sent(body, "表内标题");
                JObject hit = SpanByCode(spans, code);
                Ok(run, "ds-span-table",
                    !headingChar || (table && painted && code.StartsWith("S_", StringComparison.Ordinal)
                    && hit != null
                    && Eq(hit, "structure", "heading_1")
                    && body.IndexOf(code, StringComparison.Ordinal) >= 0),
                    table ? (Dump(spans) + " body=" + Trim(body)) : "插表失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApply(TestRun run, WdSession s)
        {
            if (Want("ds-apply-both-para", "ds-apply-both-text")) RunApplyClear(run, s);
            if (Want("ds-apply-body-text", "ds-apply-body-para")) RunApplyBody(run, s);
            if (Want("ds-apply-para-whole", "ds-apply-para-text")) RunApplyList(run, s);
            if (Want("ds-apply-char-para", "ds-apply-char-text")) RunApplyLink(run, s);
            if (Want("ds-apply-miss-text", "ds-apply-miss-index", "ds-apply-unknown", "ds-apply-one-code")) RunApplyMiss(run, s);
            if (Want("ds-apply-items-partial")) RunApplyItems(run, s);
            if (Want("ds-apply-scope-text", "ds-apply-scope-para", "ds-apply-para-ignores-scope", "ds-apply-alias")) RunApplyScope(run, s);
            if (Want("ds-apply-no-copy", "ds-apply-extra-format")) RunApplyCopy(run, s);
        }

        private static void RunApplyClear(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "整段套用。\r只套好字", out path, out id))
            {
                FailMany(run, "铺套用稿失败", "ds-apply-both-para", "ds-apply-both-text");
                return;
            }

            try
            {
                PaintParagraphSpace(s, path, "整段套用。", 18f);
                PaintBold(s, path, "整段套用。", "整");
                string body = Content(s, id);
                string whole = Sent(body, "整段套用");
                JObject applied = One(s, id, Map("action", "apply", "target_codes", whole, "structure", "heading_1", "scope", "paragraph"));
                bool styled = StyleHas(s, path, "整段套用。", HeadingNames);
                bool bold = Bold(s, path, "整段套用。", "整");
                bool paraOk = styled && (s.IsWps ? bold : !bold);
                Ok(run, "ds-apply-both-para", OkFlag(applied) && paraOk,
                    OkFlag(applied) ? "styled=" + styled + " bold=" + bold : Dump(applied));

                PaintBold(s, path, "只套好字", "好");
                body = Content(s, id);
                string part = Sent(body, "只套好字");
                JObject textApply = One(s, id, Map(
                    "action", "apply", "target_codes", part, "structure", "heading_1", "text", "好"));
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                bool textOk = Blob(textApply, "段落格式没有发生作用") && !StyleHas(s, path, "只套好字", HeadingNames) && (headingChar
                    ? !Bold(s, path, "只套好字", "好") && SliceHas(s, path, "只套好字", "好", HeadingCharNames)
                    : Eq(textApply, "scope", "text") && Bold(s, path, "只套好字", "好"));
                Ok(run, "ds-apply-both-text", OkFlag(textApply) && textOk, Dump(textApply));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyBody(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "正文几个字\r整段回正文。", out path, out id))
            {
                FailMany(run, "铺正文套用失败", "ds-apply-body-text", "ds-apply-body-para");
                return;
            }

            try
            {
                PaintSlice(s, path, "正文几个字", "几个", HeadingCharNames);
                string body = Content(s, id);
                JObject textApply = One(s, id, Map("action", "apply", "target_codes", Sent(body, "正文几个字"), "structure", "body", "text", "几个"));
                bool cleared = Blob(textApply, "段落格式没有发生作用")
                    && !SliceHas(s, path, "正文几个字", "几个", HeadingCharNames)
                    && !StyleHas(s, path, "正文几个字", HeadingNames);
                Ok(run, "ds-apply-body-text", OkFlag(textApply) && cleared, Dump(textApply));

                PaintBold(s, path, "整段回正文。", "整");
                body = Content(s, id);
                JObject paraApply = One(s, id, Map("action", "apply", "target_codes", Sent(body, "整段回正文"), "structure", "body", "scope", "paragraph"));
                Ok(run, "ds-apply-body-para", OkFlag(paraApply) && !Bold(s, path, "整段回正文。", "整"), Dump(paraApply));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyList(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "列表整段套。\r列表几个字", out path, out id))
            {
                FailMany(run, "铺列表套用失败", "ds-apply-para-whole", "ds-apply-para-text");
                return;
            }

            try
            {
                bool listBuiltin = HostHasStyle(s, path, ListNames);
                PaintSlice(s, path, "列表整段套。", "整段", HeadingCharNames);
                string body = Content(s, id);
                JObject whole = One(s, id, Map("action", "apply", "target_codes", Sent(body, "列表整段套"), "structure", "列表段落", "scope", "paragraph"));
                Ok(run, "ds-apply-para-whole",
                    listBuiltin
                        ? OkFlag(whole) && StyleHas(s, path, "列表整段套。", ListNames) && SliceHas(s, path, "列表整段套。", "整段", HeadingCharNames)
                        : !OkFlag(whole) && Blob(whole, "没有样式"),
                    Dump(whole));

                PaintSlice(s, path, "列表几个字", "几", HeadingCharNames);
                body = Content(s, id);
                JObject part = One(s, id, Map("action", "apply", "target_codes", Sent(body, "列表几个字"), "structure", "列表段落", "scope", "text", "text", "几"));
                Ok(run, "ds-apply-para-text",
                    listBuiltin
                        ? OkFlag(part) && Blob(part, "段落格式没有发生作用") && Blob(part, "什么也不会发生")
                            && !StyleHas(s, path, "列表几个字", ListNames)
                            && SliceHas(s, path, "列表几个字", "几", HeadingCharNames)
                        : !OkFlag(part) && Blob(part, "没有样式"),
                    Dump(part));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyLink(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "超链接整段。\r超链接好字", out path, out id))
            {
                FailMany(run, "铺超链接套用失败", "ds-apply-char-para", "ds-apply-char-text");
                return;
            }

            try
            {
                Materialize(s, path);
                string body = Content(s, id);
                JObject whole = One(s, id, Map("action", "apply", "target_codes", Sent(body, "超链接整段"), "structure", "超链接", "scope", "paragraph"));
                Ok(run, "ds-apply-char-para",
                    OkFlag(whole) && !StyleHas(s, path, "超链接整段。", LinkNames) && SliceHas(s, path, "超链接整段。", "整段", LinkNames),
                    Dump(whole));
                body = Content(s, id);
                JObject part = One(s, id, Map("action", "apply", "target_codes", Sent(body, "超链接好字"), "structure", "超链接", "text", "好"));
                Ok(run, "ds-apply-char-text",
                    OkFlag(part) && SliceHas(s, path, "超链接好字", "好", LinkNames) && !SliceHas(s, path, "超链接好字", "字", LinkNames),
                    Dump(part));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyMiss(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "对不上甲。", out path, out id))
            {
                FailMany(run, "铺失败套用稿失败", "ds-apply-miss-text", "ds-apply-miss-index", "ds-apply-unknown", "ds-apply-one-code");
                return;
            }

            try
            {
                string body = Content(s, id);
                string code = Sent(body, "对不上甲");
                string before = FullXml(s, path);
                JObject miss = One(s, id, Map("action", "apply", "target_codes", code, "structure", "heading_1", "text", "没有这几个字"));
                Ok(run, "ds-apply-miss-text", !OkFlag(miss) && OnlyBuiltinStyleFill(before, FullXml(s, path)), Dump(miss));
                JObject index = One(s, id, Map(
                    "action", "apply", "target_codes", code, "structure", "heading_1",
                    "text", Map("value", "甲", "index", 9)));
                Ok(run, "ds-apply-miss-index", !OkFlag(index) && OnlyBuiltinStyleFill(before, FullXml(s, path)), Dump(index));
                JObject unknown = One(s, id, Map("action", "apply", "target_codes", code, "structure", "没有这种样式"));
                Ok(run, "ds-apply-unknown", !OkFlag(unknown) && OnlyBuiltinStyleFill(before, FullXml(s, path)), Dump(unknown));
                JObject two = One(s, id, Map("action", "apply", "target_codes", code + ",S_abc12", "structure", "heading_1"));
                Ok(run, "ds-apply-one-code", !OkFlag(two), Dump(two));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyItems(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "第一项不动。\r第二项套上。", out path, out id))
            {
                FailMany(run, "铺多项稿失败", "ds-apply-items-partial");
                return;
            }

            try
            {
                string body = Content(s, id);
                JObject data = Act(s, id, Map(
                    "action", "apply",
                    "items", new List<Dictionary<string, object>>
                    {
                        Map("target_codes", Sent(body, "第一项不动"), "structure", "heading_1", "text", "没有"),
                        Map("target_codes", Sent(body, "第二项套上"), "structure", "heading_1", "scope", "paragraph")
                    }));
                JArray applied = Arr(data, "applied");
                bool firstFail = applied.Count >= 1 && !OkFlag(applied[0] as JObject);
                bool secondOk = applied.Count >= 2 && OkFlag(applied[1] as JObject) && StyleHas(s, path, "第二项套上。", HeadingNames);
                Ok(run, "ds-apply-items-partial", firstFail && secondOk && !StyleHas(s, path, "第一项不动。", HeadingNames), Dump(data));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyScope(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "省略范围好。\r省略整段。\r段落忽略范围", out path, out id))
            {
                FailMany(run, "铺范围稿失败", "ds-apply-scope-text", "ds-apply-scope-para", "ds-apply-para-ignores-scope", "ds-apply-alias");
                return;
            }

            try
            {
                string body = Content(s, id);
                JObject text = One(s, id, Map("action", "apply", "target_codes", Sent(body, "省略范围好"), "structure", "heading_1", "text", "好"));
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                Ok(run, "ds-apply-scope-text",
                    OkFlag(text) && Blob(text, "段落格式没有发生作用") && (headingChar
                        ? !StyleHas(s, path, "省略范围好。", HeadingNames) && SliceHas(s, path, "省略范围好。", "好", HeadingCharNames)
                        : Eq(text, "scope", "text") && !StyleHas(s, path, "省略范围好。", HeadingNames)),
                    Dump(text));
                body = Content(s, id);
                JObject para = One(s, id, Map("action", "apply", "target_codes", Sent(body, "省略整段"), "structure", "heading_1"));
                Ok(run, "ds-apply-scope-para", OkFlag(para) && StyleHas(s, path, "省略整段。", HeadingNames), Dump(para));
                body = Content(s, id);
                JObject ignore = One(s, id, Map("action", "apply", "target_codes", Sent(body, "段落忽略范围"), "structure", "列表段落", "scope", "text", "text", "忽略"));
                bool listBuiltin = HostHasStyle(s, path, ListNames);
                Ok(run, "ds-apply-para-ignores-scope",
                    listBuiltin
                        ? OkFlag(ignore) && Blob(ignore, "段落格式没有发生作用") && Blob(ignore, "什么也不会发生")
                            && !StyleHas(s, path, "段落忽略范围", ListNames)
                        : !OkFlag(ignore) && Blob(ignore, "没有样式"),
                    Dump(ignore));
            }
            finally
            {
                s.CloseDocByPath(path);
            }

            if (!Open(s, "别名甲。\r别名乙。\r别名丙。", out path, out id))
            {
                FailMany(run, "铺别名稿失败", "ds-apply-alias");
                return;
            }

            try
            {
                string body = Content(s, id);
                bool a = OkFlag(One(s, id, Map("action", "apply", "target_codes", Sent(body, "别名甲"), "structure", "heading_1", "scope", "paragraph")));
                bool b = OkFlag(One(s, id, Map("action", "apply", "target_codes", Sent(body, "别名乙"), "structure", "标题 1", "scope", "paragraph")));
                bool c = OkFlag(One(s, id, Map("action", "apply", "target_codes", Sent(body, "别名丙"), "structure", "heading 1", "scope", "paragraph")));
                Ok(run, "ds-apply-alias",
                    a && b && c
                    && StyleHas(s, path, "别名甲。", HeadingNames)
                    && StyleHas(s, path, "别名乙。", HeadingNames)
                    && StyleHas(s, path, "别名丙。", HeadingNames),
                    a + "," + b + "," + c);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunApplyCopy(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "不抄字号好", out path, out id))
            {
                FailMany(run, "铺不抄字号稿失败", "ds-apply-no-copy", "ds-apply-extra-format");
                return;
            }

            try
            {
                string body = Content(s, id);
                string code = Sent(body, "不抄字号好");
                JObject extra = One(s, id, Map(
                    "action", "apply", "target_codes", code, "structure", "heading_1", "text", "好",
                    "text_format", Map("font_size", 8)));
                string xml = FullXml(s, path);
                int at = xml.IndexOf("好", StringComparison.Ordinal);
                string window = at < 0 ? "" : xml.Substring(Math.Max(0, at - 400), Math.Min(400, at));
                bool copied = window.IndexOf("<w:sz", StringComparison.Ordinal) >= 0 || window.IndexOf(":sz ", StringComparison.Ordinal) >= 0;
                bool headingChar = HostHasStyle(s, path, HeadingCharNames);
                Ok(run, "ds-apply-no-copy", OkFlag(extra) && copied == !headingChar, copied ? window : Dump(extra));
                Ok(run, "ds-apply-extra-format", OkFlag(extra) && Blob(extra, "format_warnings"), Dump(extra));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEdit(TestRun run, WdSession s)
        {
            if (Want("ds-edit-heading-size")) RunEditHeading(run, s);
            if (Want("ds-edit-para-side", "ds-edit-char-side")) RunEditSides(run, s);
            if (Want("ds-edit-create-noname", "ds-edit-create-dup", "ds-edit-create-keys", "ds-edit-create-outline", "ds-edit-create-clean")) RunEditCreate(run, s);
            if (Want("ds-edit-outline-keep")) RunEditOutline(run, s);
            if (Want("ds-edit-body")) RunEditBody(run, s);
        }

        private static void RunMore(TestRun run, WdSession s)
        {
            if (Want(
                "ds-more-h3-outline", "ds-more-h9-outline", "ds-more-one-heading",
                "ds-more-action-bad", "ds-more-emphasis-plain", "ds-more-list-stable"))
            {
                RunMoreCatalog(run, s);
            }

            if (Want(
                "ds-more-missing-structure", "ds-more-missing-code", "ds-more-empty-items",
                "ds-more-apply-second", "ds-more-apply-h2", "ds-more-apply-h9",
                "ds-more-apply-title", "ds-more-apply-subtitle", "ds-more-apply-index",
                "ds-more-apply-body-keeps", "ds-more-apply-text-keeps", "ds-more-apply-char-para",
                "ds-more-items-two", "ds-more-apply-warn"))
            {
                RunMoreApply(run, s);
            }

            if (Want("ds-more-span-two", "ds-more-span-h2", "ds-more-span-after"))
            {
                RunMoreSpans(run, s);
            }

            if (Want(
                "ds-more-edit-unknown", "ds-more-edit-no-structure", "ds-more-h1-not-h2",
                "ds-more-create-para", "ds-more-create-none", "ds-more-create-struct", "ds-more-create-apply"))
            {
                RunMoreEdit(run, s);
            }
        }

        private static void RunMoreCatalog(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "ds-more-h3-outline", "ds-more-h9-outline", "ds-more-one-heading",
                "ds-more-action-bad", "ds-more-emphasis-plain", "ds-more-list-stable"
            };
            string path, id;
            if (!Open(s, "巩固目录。", out path, out id))
            {
                FailMany(run, "铺巩固目录稿失败", ids);
                return;
            }

            try
            {
                JObject listed = Act(s, id, "list_styles");
                JArray styles = Arr(listed, "styles");
                JObject heading3 = Find(styles, "heading_3");
                JObject heading9 = Find(styles, "heading_9");
                int headingCount = 0;
                foreach (JObject row in Each(styles))
                {
                    if (Eq(row, "structure", "heading_1"))
                    {
                        headingCount++;
                    }
                }

                Ok(run, "ds-more-h3-outline",
                    heading3 != null && Eq(Obj(heading3, "paragraph_format"), "outline_level", "wdOutlineLevel3"),
                    Dump(heading3));
                Ok(run, "ds-more-h9-outline",
                    heading9 != null && Eq(Obj(heading9, "paragraph_format"), "outline_level", "wdOutlineLevel9"),
                    Dump(heading9));
                Ok(run, "ds-more-one-heading", headingCount == 1, "heading_1=" + headingCount);
                JObject bad = Act(s, id, "list_foo");
                Ok(run, "ds-more-action-bad", !OkFlag(bad) && Blob(bad, "action 须是"), Dump(bad));
                JObject emphasis = Find(styles, "明显强调");
                bool emphasisBuiltin = HostHasStyle(s, path, "明显强调", "Intense Emphasis");
                Ok(run, "ds-more-emphasis-plain",
                    emphasisBuiltin
                        ? emphasis != null && Eq(emphasis, "support", "character") && Obj(emphasis, "paragraph_format") == null
                        : emphasis == null,
                    Dump(emphasis));
                string support = Str(Find(styles, "heading_1"), "support");
                JObject again = Act(s, id, "list_styles");
                JArray second = Arr(again, "styles");
                Ok(run, "ds-more-list-stable",
                    support.Length > 0 && support == Str(Find(second, "heading_1"), "support") && styles.Count == second.Count,
                    support + " / " + second.Count);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunMoreApply(TestRun run, WdSession s)
        {
            string[] ids =
            {
                "ds-more-missing-structure", "ds-more-missing-code", "ds-more-empty-items",
                "ds-more-apply-second", "ds-more-apply-h2", "ds-more-apply-h9",
                "ds-more-apply-title", "ds-more-apply-subtitle", "ds-more-apply-index",
                "ds-more-apply-body-keeps", "ds-more-apply-text-keeps", "ds-more-apply-char-para",
                "ds-more-items-two", "ds-more-apply-warn"
            };
            string path, id;
            if (!Open(s, "第一句甲。\r第二句乙。\r二级这句。\r九级这句。\r标题这句。\r副题这句。\r甲甲字。\r甲段。\r乙段。\r只套好字。\r整段链。\r丙句。\r丁句。", out path, out id))
            {
                FailMany(run, "铺巩固套用稿失败", ids);
                return;
            }

            try
            {
                Materialize(s, path);
                string body = Content(s, id);
                string first = Sent(body, "第一句甲");
                string second = Sent(body, "第二句乙");
                string level2 = Sent(body, "二级这句");
                string level9 = Sent(body, "九级这句");
                string title = Sent(body, "标题这句");
                string subtitle = Sent(body, "副题这句");
                string repeat = Sent(body, "甲甲字");
                string keepA = Sent(body, "甲段");
                string slice = Sent(body, "只套好字");
                string link = Sent(body, "整段链");
                string itemC = Sent(body, "丙句");
                string itemD = Sent(body, "丁句");
                JObject missStructure = One(s, id, Map("action", "apply", "target_codes", first));
                Ok(run, "ds-more-missing-structure",
                    !OkFlag(missStructure) && Blob(missStructure, "structure") && !StyleHas(s, path, "第一句甲", HeadingNames),
                    Dump(missStructure));
                JObject missCode = One(s, id, Map("action", "apply", "structure", "heading_1"));
                Ok(run, "ds-more-missing-code", !OkFlag(missCode) && Blob(missCode, "target_codes"), Dump(missCode));
                JObject empty = Act(s, id, Map("action", "apply", "items", new List<Dictionary<string, object>>()));
                Ok(run, "ds-more-empty-items", !OkFlag(empty) && Blob(empty, "没有套上"), Dump(empty));
                JObject secondApply = One(s, id, Map("action", "apply", "target_codes", second, "structure", "heading_1"));
                Ok(run, "ds-more-apply-second",
                    OkFlag(secondApply) && StyleHas(s, path, "第二句乙", HeadingNames) && !StyleHas(s, path, "第一句甲", HeadingNames),
                    Dump(secondApply));
                JObject h2 = One(s, id, Map("action", "apply", "target_codes", level2, "structure", "heading_2"));
                Ok(run, "ds-more-apply-h2",
                    OkFlag(h2) && StyleHas(s, path, "二级这句", Heading2Names) && !StyleHas(s, path, "二级这句", HeadingNames),
                    Dump(h2));
                JObject h9 = One(s, id, Map("action", "apply", "target_codes", level9, "structure", "heading_9"));
                Ok(run, "ds-more-apply-h9", OkFlag(h9) && StyleHas(s, path, "九级这句", Heading9Names), Dump(h9));
                JObject titleApply = One(s, id, Map("action", "apply", "target_codes", title, "structure", "title"));
                Ok(run, "ds-more-apply-title",
                    OkFlag(titleApply) && StyleIs(s, path, "标题这句", TitleNames) && !StyleHas(s, path, "标题这句", HeadingNames),
                    Dump(titleApply));
                JObject subtitleApply = One(s, id, Map("action", "apply", "target_codes", subtitle, "structure", "subtitle"));
                Ok(run, "ds-more-apply-subtitle",
                    OkFlag(subtitleApply) && StyleIs(s, path, "副题这句", SubtitleNames) && !StyleHas(s, path, "副题这句", Heading2Names),
                    Dump(subtitleApply));
                JObject indexed = One(s, id, Map(
                    "action", "apply", "target_codes", repeat, "structure", "超链接",
                    "text", Map("value", "甲", "index", 1)));
                Ok(run, "ds-more-apply-index",
                    OkFlag(indexed) && SliceNthHas(s, path, "甲甲字", "甲", 1, LinkNames) && !SliceNthHas(s, path, "甲甲字", "甲", 2, LinkNames),
                    Dump(indexed));
                bool painted = PaintParagraph(s, path, "乙段。", HeadingNames);
                JObject bodyApply = One(s, id, Map("action", "apply", "target_codes", keepA, "structure", "body"));
                Ok(run, "ds-more-apply-body-keeps",
                    painted && OkFlag(bodyApply) && StyleHas(s, path, "甲段", BodyNames) && StyleHas(s, path, "乙段", HeadingNames),
                    Dump(bodyApply));
                JObject textApply = One(s, id, Map(
                    "action", "apply", "target_codes", slice, "structure", "heading_1", "text", "好"));
                Ok(run, "ds-more-apply-text-keeps",
                    OkFlag(textApply) && Blob(textApply, "段落格式没有发生作用")
                    && StyleHas(s, path, "只套好字", BodyNames) && !StyleHas(s, path, "只套好字", HeadingNames),
                    Dump(textApply));
                JObject charApply = One(s, id, Map(
                    "action", "apply", "target_codes", link, "structure", "超链接", "scope", "paragraph"));
                Ok(run, "ds-more-apply-char-para",
                    OkFlag(charApply) && SliceHas(s, path, "整段链", "整", LinkNames)
                    && StyleHas(s, path, "整段链", BodyNames) && !StyleHas(s, path, "整段链", LinkNames),
                    Dump(charApply));
                JObject both = Act(s, id, Map(
                    "action", "apply",
                    "items", new List<Dictionary<string, object>>
                    {
                        Map("target_codes", itemC, "structure", "heading_1"),
                        Map("target_codes", itemD, "structure", "heading_2")
                    }));
                Ok(run, "ds-more-items-two",
                    OkFlag(both) && StyleHas(s, path, "丙句", HeadingNames) && StyleHas(s, path, "丁句", Heading2Names),
                    Dump(both));
                JObject warn = One(s, id, Map(
                    "action", "apply", "target_codes", first, "structure", "heading_1",
                    "paragraph_format", Map("space_before", 40)));
                Ok(run, "ds-more-apply-warn",
                    OkFlag(warn) && Blob(warn, "format_warnings") && StyleHas(s, path, "第一句甲", HeadingNames),
                    Dump(warn));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunMoreSpans(TestRun run, WdSession s)
        {
            string[] ids = { "ds-more-span-two", "ds-more-span-h2", "ds-more-span-after" };
            string path, id;
            if (!Open(s, "甲句。\r乙句。", out path, out id))
            {
                FailMany(run, "铺巩固分截稿失败", ids);
                return;
            }

            try
            {
                string body = Content(s, id);
                string codeA = Sent(body, "甲句");
                string codeB = Sent(body, "乙句");
                JArray spans = Arr(Act(s, id, "list_spans"), "spans");
                JObject spanA = SpanByCode(spans, codeA);
                JObject spanB = SpanByCode(spans, codeB);
                Ok(run, "ds-more-span-two",
                    spanA != null && spanB != null && codeA != codeB
                    && Eq(spanA, "structure", "body") && Eq(spanB, "structure", "body")
                    && Eq(spanA, "scope", "paragraph") && Eq(spanB, "scope", "paragraph"),
                    Dump(spanA) + " / " + Dump(spanB));
                JObject h2 = One(s, id, Map("action", "apply", "target_codes", codeB, "structure", "heading_2"));
                JObject afterH2 = SpanByCode(Arr(Act(s, id, "list_spans"), "spans"), codeB);
                Ok(run, "ds-more-span-h2",
                    OkFlag(h2) && afterH2 != null && Eq(afterH2, "structure", "heading_2") && Eq(afterH2, "scope", "paragraph"),
                    Dump(afterH2));
                JObject h1 = One(s, id, Map("action", "apply", "target_codes", codeA, "structure", "heading_1"));
                JArray later = Arr(Act(s, id, "list_spans"), "spans");
                JObject afterH1 = SpanByCode(later, codeA);
                JObject stillH2 = SpanByCode(later, codeB);
                Ok(run, "ds-more-span-after",
                    OkFlag(h1) && afterH1 != null && Eq(afterH1, "structure", "heading_1")
                    && stillH2 != null && Eq(stillH2, "structure", "heading_2"),
                    Dump(afterH1) + " / " + Dump(stillH2));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunMoreEdit(TestRun run, WdSession s)
        {
            string[] ids = { "ds-more-edit-unknown", "ds-more-edit-no-structure", "ds-more-h1-not-h2" };
            string path, id;
            if (!Open(s, "改一不改二。", out path, out id))
            {
                FailMany(run, "铺巩固改稿失败", ids);
                return;
            }

            try
            {
                JObject unknown = Act(s, id, Map("action", "customize", "structure", "没有这种样式", "text_format", Map("font_size", 16)));
                Ok(run, "ds-more-edit-unknown", !OkFlag(unknown) && Blob(unknown, "没有样式"), Dump(unknown));
                JObject missing = Act(s, id, Map("action", "customize", "text_format", Map("font_size", 16)));
                Ok(run, "ds-more-edit-no-structure", !OkFlag(missing) && Blob(missing, "structure"), Dump(missing));
                JObject edited = Act(s, id, Map("action", "customize", "structure", "heading_1", "text_format", Map("font_size", 19)));
                JArray styles = Arr(Act(s, id, "list_styles"), "styles");
                JObject heading1 = Find(styles, "heading_1");
                JObject heading2 = Find(styles, "heading_2");
                Ok(run, "ds-more-h1-not-h2",
                    OkFlag(edited) && heading2 != null
                    && Near(Obj(heading1, "text_format"), "font_size", 19, 0.6)
                    && !Near(Obj(heading2, "text_format"), "font_size", 19, 0.6),
                    Dump(heading1) + " / " + Dump(heading2));
            }
            finally
            {
                s.CloseDocByPath(path);
            }

            string[] createIds = { "ds-more-create-para", "ds-more-create-none", "ds-more-create-struct", "ds-more-create-apply" };
            if (!Want(createIds))
            {
                return;
            }

            if (!Open(s, "套上用。", out path, out id))
            {
                FailMany(run, "铺巩固新建稿失败", createIds);
                return;
            }

            try
            {
                JObject withStructure = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "巩固坏", "structure", "heading_1",
                    "paragraph_format", Map("left_indent", 12)));
                JObject para = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "巩固段",
                    "paragraph_format", Map("left_indent", 12)));
                JObject none = Act(s, id, Map("action", "customize", "create", true, "name", "巩固空"));
                JObject created = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "巩固用",
                    "text_format", Map("font_size", 15)));
                JArray styles = Arr(Act(s, id, "list_styles"), "styles");
                JObject paraRow = Find(styles, "巩固段");
                JObject noneRow = Find(styles, "巩固空");
                JObject useRow = Find(styles, "巩固用");
                Ok(run, "ds-more-create-struct",
                    !OkFlag(withStructure) && Blob(withStructure, "structure") && Find(styles, "巩固坏") == null,
                    Dump(withStructure));
                Ok(run, "ds-more-create-para",
                    OkFlag(para) && paraRow != null && Eq(paraRow, "support", "paragraph")
                    && Obj(paraRow, "text_format") == null
                    && Near(Obj(paraRow, "paragraph_format"), "left_indent", 12, 1.2)
                    && Find(styles, "巩固段 字符") == null,
                    Dump(paraRow));
                Ok(run, "ds-more-create-none",
                    OkFlag(none) && noneRow != null && Eq(noneRow, "support", "paragraph") && Find(styles, "巩固空 字符") == null,
                    Dump(noneRow));
                string body = Content(s, id);
                string code = Sent(body, "套上用");
                JObject applied = One(s, id, Map(
                    "action", "apply", "target_codes", code, "structure", "巩固用", "text", "用"));
                Ok(run, "ds-more-create-apply",
                    OkFlag(created) && useRow != null && Eq(useRow, "support", "character")
                    && Near(Obj(useRow, "text_format"), "font_size", 15, 0.6)
                    && Obj(useRow, "paragraph_format") == null
                    && OkFlag(applied) && SliceHas(s, path, "套上用", "用", new[] { "巩固用" }),
                    Dump(useRow) + Dump(applied));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEditHeading(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "跟着变字。\r盖住字号。", out path, out id))
            {
                FailMany(run, "铺改字号稿失败", "ds-edit-heading-size");
                return;
            }

            try
            {
                PaintParagraph(s, path, "跟着变字。", HeadingNames);
                PaintParagraph(s, path, "盖住字号。", HeadingNames);
                PaintSize(s, path, "盖住字号。", "盖", 28f);
                JObject edited = Act(s, id, Map(
                    "action", "customize", "structure", "heading_1",
                    "text_format", Map("font_size", 20)));
                float follow = ReadSize(s, path, "跟着变字。", "跟");
                float direct = ReadSize(s, path, "盖住字号。", "盖");
                float paraSize = StyleSize(s, path, HeadingNames);
                float charSize = StyleSize(s, path, HeadingCharNames);
                bool linked = StylesLinked(s, path, HeadingNames, HeadingCharNames);
                Ok(run, "ds-edit-heading-size",
                    OkFlag(edited) && Math.Abs(follow - 20f) < 0.6f && Math.Abs(direct - 28f) < 0.6f
                    && Math.Abs(paraSize - 20f) < 0.6f && Math.Abs(charSize - 20f) < 0.6f && linked,
                    "follow=" + follow.ToString(CultureInfo.InvariantCulture)
                    + " direct=" + direct.ToString(CultureInfo.InvariantCulture)
                    + " para=" + paraSize.ToString(CultureInfo.InvariantCulture)
                    + " char=" + charSize.ToString(CultureInfo.InvariantCulture)
                    + " linked=" + linked + " " + Dump(edited));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEditSides(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "改两侧。", out path, out id))
            {
                FailMany(run, "铺改两侧稿失败", "ds-edit-para-side", "ds-edit-char-side");
                return;
            }

            try
            {
                Materialize(s, path);
                bool listBuiltin = HostHasStyle(s, path, ListNames);
                JObject para = Act(s, id, Map(
                    "action", "customize", "structure", "列表段落",
                    "paragraph_format", Map("left_indent", 24),
                    "text_format", Map("font_size", 9)));
                JObject styles = Act(s, id, "list_styles");
                JObject list = Find(Arr(styles, "styles"), "列表段落");
                Ok(run, "ds-edit-para-side",
                    listBuiltin
                        ? OkFlag(para) && list != null && Eq(list, "support", "both")
                            && Near(Obj(list, "paragraph_format"), "left_indent", 24, 1.2)
                            && Near(Obj(list, "text_format"), "font_size", 9, 0.6)
                        : !OkFlag(para) && Blob(para, "没有样式") && list == null,
                    Dump(para) + Dump(list));
                JObject character = Act(s, id, Map(
                    "action", "customize", "structure", "超链接",
                    "text_format", Map("font_color", "FF00FF"),
                    "paragraph_format", Map("outline_level", "wdOutlineLevel2")));
                JObject link = Find(Arr(Act(s, id, "list_styles"), "styles"), "超链接");
                string linkColor = Str(Obj(link, "text_format"), "font_color");
                Ok(run, "ds-edit-char-side",
                    Blob(character, "format_warnings") && linkColor.IndexOf("FF00FF", StringComparison.OrdinalIgnoreCase) >= 0
                    && (Obj(link, "paragraph_format") == null || Str(Obj(link, "paragraph_format"), "outline_level").Length == 0),
                    Dump(character) + Dump(link));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEditCreate(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "新建样式。", out path, out id))
            {
                FailMany(run, "铺新建稿失败", "ds-edit-create-noname", "ds-edit-create-dup", "ds-edit-create-keys", "ds-edit-create-outline", "ds-edit-create-clean");
                return;
            }

            try
            {
                JObject noname = Act(s, id, Map("action", "customize", "create", true));
                Ok(run, "ds-edit-create-noname", !OkFlag(noname), Dump(noname));
                JObject dup = Act(s, id, Map("action", "customize", "create", true, "name", "标题 1"));
                JObject heading = Find(Arr(Act(s, id, "list_styles"), "styles"), "heading_1");
                Ok(run, "ds-edit-create-dup", !OkFlag(dup) && heading != null, Dump(dup));
                JObject created = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "强调测",
                    "text_format", Map("font_size", 16)));
                JObject row = Find(Arr(Act(s, id, "list_styles"), "styles"), "强调测");
                bool sizeOnly = row != null && Eq(row, "support", "character")
                    && Near(Obj(row, "text_format"), "font_size", 16, 0.6)
                    && Obj(row, "paragraph_format") == null;
                Ok(run, "ds-edit-create-keys", OkFlag(created) && sizeOnly, Dump(created) + Dump(row));
                JObject outlined = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "强调纲",
                    "paragraph_format", Map("outline_level", "wdOutlineLevel2")));
                JObject outlineRow = Find(Arr(Act(s, id, "list_styles"), "styles"), "强调纲");
                Ok(run, "ds-edit-create-outline",
                    OkFlag(outlined) && outlineRow != null && Eq(outlineRow, "structure", "强调纲")
                    && Eq(outlineRow, "support", "paragraph")
                    && Obj(outlineRow, "text_format") == null
                    && Str(Obj(outlineRow, "paragraph_format"), "outline_level").IndexOf("2", StringComparison.Ordinal) >= 0,
                    Dump(outlined) + Dump(outlineRow));
                JObject dual = Act(s, id, Map(
                    "action", "customize", "create", true, "name", "强调双",
                    "text_format", Map("font_size", 16),
                    "paragraph_format", Map("left_indent", 18)));
                JObject dualRow = Find(Arr(Act(s, id, "list_styles"), "styles"), "强调双");
                float dualPara = StyleSize(s, path, new[] { "强调双" });
                float dualChar = StyleSize(s, path, new[] { "强调双 字符" });
                bool dualLinked = StylesLinked(s, path, new[] { "强调双" }, new[] { "强调双 字符" });
                Ok(run, "ds-edit-create-clean",
                    OkFlag(dual) && dualRow != null && Eq(dualRow, "support", "both")
                    && Find(Arr(Act(s, id, "list_styles"), "styles"), "强调双 字符") == null
                    && Near(Obj(dualRow, "text_format"), "font_size", 16, 0.6)
                    && Near(Obj(dualRow, "paragraph_format"), "left_indent", 18, 1.2)
                    && Math.Abs(dualPara - 16f) < 0.6f && Math.Abs(dualChar - 16f) < 0.6f && dualLinked,
                    "para=" + dualPara.ToString(CultureInfo.InvariantCulture)
                    + " char=" + dualChar.ToString(CultureInfo.InvariantCulture)
                    + " linked=" + dualLinked + " " + Dump(dual) + Dump(dualRow));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEditOutline(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "改大纲。", out path, out id))
            {
                FailMany(run, "铺大纲稿失败", "ds-edit-outline-keep");
                return;
            }

            try
            {
                JObject edited = Act(s, id, Map(
                    "action", "customize", "structure", "heading_1",
                    "paragraph_format", Map("outline_level", "wdOutlineLevel3")));
                s.OnWordDoc(path, doc =>
                {
                    doc.Save();
                    return true;
                });
                JObject heading = Find(Arr(Act(s, id, "list_styles"), "styles"), "heading_1");
                int outline = ReadOutline(s, path, HeadingNames);
                bool kept = heading != null && Eq(heading, "structure", "heading_1")
                    && Eq(Obj(heading, "paragraph_format"), "outline_level", "wdOutlineLevel1");
                Ok(run, "ds-edit-outline-keep",
                    !OkFlag(edited) && Blob(edited, "不允许改大纲层级") && kept && outline == 1,
                    "com=" + outline + " " + Dump(edited) + Dump(heading));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunEditBody(TestRun run, WdSession s)
        {
            string path, id;
            if (!Open(s, "正文跟着。\r标题不跟。", out path, out id))
            {
                FailMany(run, "铺正文稿失败", "ds-edit-body");
                return;
            }

            try
            {
                PaintParagraph(s, path, "标题不跟。", HeadingNames);
                float headingBefore = ReadSize(s, path, "标题不跟。", "标");
                JObject edited = Act(s, id, Map(
                    "action", "customize", "structure", "body",
                    "text_format", Map("font_size", 12)));
                float bodySize = ReadSize(s, path, "正文跟着。", "正");
                float headingAfter = ReadSize(s, path, "标题不跟。", "标");
                float bodyStyle = StyleSize(s, path, BodyNames);
                float bodyChar = StyleSize(s, path, new[] { "正文 字符" });
                bool bodyLinked = StylesLinked(s, path, BodyNames, new[] { "正文 字符" });
                Ok(run, "ds-edit-body",
                    OkFlag(edited) && Math.Abs(bodySize - 12f) < 0.6f && Math.Abs(headingAfter - headingBefore) < 0.6f && Math.Abs(headingAfter - 12f) > 0.6f
                    && Math.Abs(bodyStyle - 12f) < 0.6f && Math.Abs(bodyChar - 12f) < 0.6f && bodyLinked,
                    "body=" + bodySize.ToString(CultureInfo.InvariantCulture)
                    + " heading=" + headingAfter.ToString(CultureInfo.InvariantCulture)
                    + " style=" + bodyStyle.ToString(CultureInfo.InvariantCulture)
                    + " char=" + bodyChar.ToString(CultureInfo.InvariantCulture)
                    + " linked=" + bodyLinked);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static bool Open(WdSession s, string plain, out string path, out string id)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            return s.TypePlain(path, plain);
        }

        private static void Materialize(WdSession s, string path)
        {
            string[] names = { "超链接", "Hyperlink", "题注", "Caption", "页眉", "Header", "页脚", "Footer", "TOC 1", "目录 1" };
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc => TouchStyles(doc, names));
            }
            else
            {
                s.OnWordDoc(path, doc => TouchStyles(doc, names));
            }
        }

        private static bool TouchStyles(object doc, string[] names)
        {
            foreach (string name in names)
            {
                try
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    WdSession.ComCall(styles, "Item", name);
                }
                catch (Exception)
                {
                }
            }

            return true;
        }

        private static bool TouchStyles(Word.Document doc, string[] names)
        {
            if (doc.Content.End > 1)
            {
                doc.Content.InsertAfter("\r样式垫。");
            }

            Word.Paragraph scratch = doc.Paragraphs.Last;
            foreach (string name in names)
            {
                try
                {
                    Word.Style style = doc.Styles[name];
                    if (style == null)
                    {
                        continue;
                    }

                    object styleRef = name;
                    if (style.Type == Word.WdStyleType.wdStyleTypeCharacter)
                    {
                        Word.Range range = scratch.Range.Duplicate;
                        if (range.End > range.Start + 1)
                        {
                            range.End = range.Start + 1;
                        }

                        range.set_Style(ref styleRef);
                    }
                    else
                    {
                        scratch.Range.set_Style(ref styleRef);
                    }
                }
                catch (Exception)
                {
                }
            }

            try
            {
                Word.HeaderFooter header = doc.Sections[1].Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                header.Range.Text = "页眉垫";
                Word.HeaderFooter footer = doc.Sections[1].Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
                footer.Range.Text = "页脚垫";
            }
            catch (Exception)
            {
            }

            return true;
        }

        private static int ReadOutline(WdSession s, string path, string[] names)
        {
            int level = 0;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    foreach (string name in names)
                    {
                        try
                        {
                            object style = WdSession.ComCall(styles, "Item", name);
                            object type = WdSession.ComGet(style, "Type");
                            if (type != null && Convert.ToInt32(type) == (int)Word.WdStyleType.wdStyleTypeCharacter)
                            {
                                continue;
                            }

                            object format = WdSession.ComGet(style, "ParagraphFormat");
                            level = Convert.ToInt32(WdSession.ComGet(format, "OutlineLevel"));
                            if (level > 0)
                            {
                                return true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return false;
                });
                return level;
            }

            s.OnWordDoc(path, doc =>
            {
                foreach (string name in names)
                {
                    try
                    {
                        Word.Style style = doc.Styles[name];
                        if (style == null || style.Type == Word.WdStyleType.wdStyleTypeCharacter)
                        {
                            continue;
                        }

                        level = (int)style.ParagraphFormat.OutlineLevel;
                        if (level > 0)
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
            return level;
        }

        private static float StyleSize(WdSession s, string path, string[] names)
        {
            float size = 0;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    foreach (string name in names)
                    {
                        try
                        {
                            object style = WdSession.ComCall(styles, "Item", name);
                            object font = WdSession.ComGet(style, "Font");
                            size = Convert.ToSingle(WdSession.ComGet(font, "Size"), CultureInfo.InvariantCulture);
                            if (size > 0 && size < 1000)
                            {
                                return true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return false;
                });
                return size;
            }

            s.OnWordDoc(path, doc =>
            {
                foreach (string name in names)
                {
                    try
                    {
                        Word.Style style = doc.Styles[name];
                        if (style == null)
                        {
                            continue;
                        }

                        size = style.Font.Size;
                        if (size > 0 && size < 1000)
                        {
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
            return size;
        }

        private static bool StylesLinked(WdSession s, string path, string[] paragraphNames, string[] characterNames)
        {
            bool linked = false;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    linked = LinkedLate(doc, paragraphNames, characterNames);
                    return true;
                });
                return linked;
            }

            s.OnWordDoc(path, doc =>
            {
                linked = LinkedWord(doc, paragraphNames, characterNames);
                return true;
            });
            return linked;
        }

        private static bool LinkedWord(Word.Document doc, string[] paragraphNames, string[] characterNames)
        {
            Word.Style paragraph = FirstStyle(doc, paragraphNames, false);
            Word.Style character = FirstStyle(doc, characterNames, true);
            if (paragraph == null || character == null)
            {
                return false;
            }

            try
            {
                Word.Style paragraphLink = paragraph.get_LinkStyle();
                Word.Style characterLink = character.get_LinkStyle();
                return paragraphLink != null && characterLink != null
                    && paragraphLink.NameLocal == character.NameLocal
                    && characterLink.NameLocal == paragraph.NameLocal;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static Word.Style FirstStyle(Word.Document doc, string[] names, bool character)
        {
            foreach (string name in names)
            {
                try
                {
                    Word.Style style = doc.Styles[name];
                    if (style == null)
                    {
                        continue;
                    }

                    bool isCharacter = style.Type == Word.WdStyleType.wdStyleTypeCharacter;
                    if (isCharacter == character)
                    {
                        return style;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static bool LinkedLate(object doc, string[] paragraphNames, string[] characterNames)
        {
            object styles = WdSession.ComGet(doc, "Styles");
            object paragraph = FirstStyleLate(styles, paragraphNames, false);
            object character = FirstStyleLate(styles, characterNames, true);
            if (paragraph == null || character == null)
            {
                return false;
            }

            try
            {
                object paragraphLink = WdSession.ComGet(paragraph, "LinkStyle");
                object characterLink = WdSession.ComGet(character, "LinkStyle");
                string paragraphLinkName = Convert.ToString(WdSession.ComGet(paragraphLink, "NameLocal"));
                string characterLinkName = Convert.ToString(WdSession.ComGet(characterLink, "NameLocal"));
                string paragraphName = Convert.ToString(WdSession.ComGet(paragraph, "NameLocal"));
                string characterName = Convert.ToString(WdSession.ComGet(character, "NameLocal"));
                return paragraphLinkName == characterName && characterLinkName == paragraphName;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static object FirstStyleLate(object styles, string[] names, bool character)
        {
            foreach (string name in names)
            {
                try
                {
                    object style = WdSession.ComCall(styles, "Item", name);
                    object type = WdSession.ComGet(style, "Type");
                    bool isCharacter = type != null && Convert.ToInt32(type) == (int)Word.WdStyleType.wdStyleTypeCharacter;
                    if (isCharacter == character)
                    {
                        return style;
                    }
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static JObject Act(WdSession s, string id, string action)
        {
            return Act(s, id, Map("action", action));
        }

        private static JObject Act(WdSession s, string id, Dictionary<string, object> args)
        {
            args["channel_id"] = id;
            ToolResult result = s.DocumentStructure(args);
            JObject data = Parse(result);
            if (data == null)
            {
                data = new JObject();
            }

            data["_success"] = result != null && result.Success;
            data["_error"] = result == null ? "" : (result.Error ?? "");
            return data;
        }

        private static JObject One(WdSession s, string id, Dictionary<string, object> args)
        {
            JObject data = Act(s, id, args);
            JArray applied = Arr(data, "applied");
            if (applied.Count == 1 && applied[0] is JObject item)
            {
                if (item["format_warnings"] == null && data["format_warnings"] != null)
                {
                    item["format_warnings"] = data["format_warnings"];
                }

                item["_success"] = OkFlag(item);
                item["_error"] = Str(item, "error");
                return item;
            }

            return data;
        }

        private static string Content(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "sentence");
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

            Match m = Regex.Match(body.Substring(mark), @"\[S_([0-9A-Za-z]{5})(?:-start| start)");
            return m.Success ? "S_" + m.Groups[1].Value : "";
        }

        private static bool PaintParagraph(WdSession s, string path, string needle, string[] names)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => PaintParagraphLate(doc, needle, names))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, needle);
                    return para != null && SetStyle(para.Range, names);
                });
        }

        private static bool PaintSlice(WdSession s, string path, string paraNeedle, string slice, string[] names)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => PaintSliceLate(doc, paraNeedle, slice, names))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    return range != null && SetStyle(range, names);
                });
        }

        private static bool PaintNth(WdSession s, string path, string paraNeedle, string slice, int nth, string[] names)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => PaintNthLate(doc, paraNeedle, slice, nth, names))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Nth(doc, paraNeedle, slice, nth);
                    return range != null && SetStyle(range, names);
                });
        }

        private static bool PaintBold(WdSession s, string path, string paraNeedle, string slice)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = SliceLate(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    WdSession.ComSet(WdSession.ComGet(range, "Font"), "Bold", 1);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    range.Font.Bold = 1;
                    return true;
                });
        }

        private static bool PaintColor(WdSession s, string path, string paraNeedle, string slice)
        {
            return s.OnWordDoc(path, doc =>
            {
                Word.Range range = Slice(doc, paraNeedle, slice);
                if (range == null)
                {
                    return false;
                }

                range.Font.Color = Word.WdColor.wdColorRed;
                return true;
            }) || s.TryWpsDoc(path, doc =>
            {
                object range = SliceLate(doc, paraNeedle, slice);
                if (range == null)
                {
                    return false;
                }

                WdSession.ComSet(WdSession.ComGet(range, "Font"), "Color", 255);
                return true;
            });
        }

        private static bool PaintSize(WdSession s, string path, string paraNeedle, string slice, float size)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = SliceLate(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    WdSession.ComSet(WdSession.ComGet(range, "Font"), "Size", size);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    range.Font.Size = size;
                    return true;
                });
        }

        private static bool PaintParagraphSpace(WdSession s, string path, string needle, float before)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => PaintSpaceLate(doc, needle, before))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, needle);
                    if (para == null)
                    {
                        return false;
                    }

                    para.Format.SpaceBefore = before;
                    return true;
                });
        }

        private static string ParagraphContent(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "paragraph");
            return got != null && got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool SentenceHolds(string body, string text, string extra)
        {
            string code = Sent(body, text);
            if (!code.StartsWith("S_", StringComparison.Ordinal))
            {
                return false;
            }

            int at = (body ?? "").IndexOf(text, StringComparison.Ordinal);
            int mark = at < 0 ? -1 : (body ?? "").LastIndexOf("[S_", at, StringComparison.Ordinal);
            if (mark < 0)
            {
                return false;
            }

            int end = (body ?? "").IndexOf("[S_" + code.Substring(2) + "-end]", mark, StringComparison.Ordinal);
            if (end < 0)
            {
                return false;
            }

            return (body ?? "").Substring(mark, end - mark).IndexOf(extra, StringComparison.Ordinal) >= 0;
        }

        private static string TocParagraph(string display)
        {
            if (string.IsNullOrEmpty(display))
            {
                return "";
            }

            foreach (Match m in Regex.Matches(display, @"\[(P_[0-9A-Za-z]{5}(?:_\d+)?) start\]([\s\S]*?)\[\1 end\]"))
            {
                string inner = m.Groups[2].Value ?? "";
                if (inner.IndexOf('\t') >= 0 && inner.IndexOf("你好啊", StringComparison.Ordinal) < 0)
                {
                    return m.Groups[1].Value;
                }
            }

            return "";
        }

        private static bool SpanHasTab(JArray spans)
        {
            foreach (JObject span in Each(spans))
            {
                string text = TextValue(span) ?? "";
                if (text.IndexOf('\t') >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool UpdateFields(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object fields = WdSession.ComGet(doc, "Fields");
                    WdSession.ComCall(fields, "Update");
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    doc.Fields.Update();
                    return true;
                });
        }

        private static int CountBookmarks(WdSession s, string path, string name)
        {
            int count = 0;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    count = CountNamedLate(doc, name);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    foreach (Word.Bookmark bookmark in doc.Bookmarks)
                    {
                        if (string.Equals(bookmark.Name, name, StringComparison.Ordinal))
                        {
                            count++;
                        }
                    }

                    return true;
                });
            return ok ? count : -1;
        }

        private static int CountNamedLate(object doc, string name)
        {
            int count = 0;
            object bookmarks = WdSession.ComGet(doc, "Bookmarks");
            if (bookmarks == null)
            {
                return 0;
            }

            foreach (object bookmark in (System.Collections.IEnumerable)bookmarks)
            {
                string found = Convert.ToString(WdSession.ComGet(bookmark, "Name"));
                if (string.Equals(found, name, StringComparison.Ordinal))
                {
                    count++;
                }
            }

            return count;
        }

        private static string PlantNote;

        private static bool PlantTocBookmark(WdSession s, string path)
        {
            PlantNote = "";
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    try
                    {
                        object range = WdSession.ComGet(doc, "Content");
                        WdSession.ComCall(range, "Collapse", 1);
                        object bookmarks = WdSession.ComGet(doc, "Bookmarks");
                        WdSession.ComCall(bookmarks, "Add", "_Toc245561084", range);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        PlantNote = ex.Message;
                        return false;
                    }
                })
                : s.OnWordDoc(path, doc =>
                {
                    try
                    {
                        Word.Range range = doc.Content.Duplicate;
                        range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        object anchor = range;
                        doc.Bookmarks.Add("_Toc245561084", ref anchor);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        PlantNote = ex.Message;
                        return false;
                    }
                });
        }

        private static bool BookmarkExists(WdSession s, string path, string name)
        {
            bool found = false;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object bookmarks = WdSession.ComGet(doc, "Bookmarks");
                    object exists = WdSession.ComCall(bookmarks, "Exists", name);
                    found = exists != null && Convert.ToBoolean(exists);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    found = doc.Bookmarks.Exists(name);
                    return true;
                });
            return ok && found;
        }

        private static string RejectInteriorSentence(WdSession s, string path)
        {
            string error = "";
            bool ok = s.OnWordDoc(path, doc =>
            {
                try
                {
                    WordDocumentExtractor.ProcessDocument(doc);
                }
                catch (Exception ex)
                {
                    error = ex.GetType().Name + ": " + ex.Message;
                    return true;
                }

                string code = TocFieldStamp.FirstInteriorSentence();
                if (string.IsNullOrEmpty(code))
                {
                    error = "没有域内句子码";
                    return true;
                }

                Dictionary<string, object> result = StructureStyleApplier.Apply(
                    doc,
                    new Dictionary<string, object>
                    {
                        ["structure"] = "标题 1",
                        ["target_codes"] = code
                    },
                    new List<StructureStyleRecord>(),
                    true);
                error = result != null && result.ContainsKey("error") ? Convert.ToString(result["error"]) : "";
                return true;
            });
            return ok ? error : "没打开稿";
        }

        private static bool AddToc(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = WdSession.ComGet(doc, "Range");
                    WdSession.ComCall(range, "Collapse", 1);
                    object fields = WdSession.ComGet(doc, "Fields");
                    WdSession.ComCall(fields, "Add", range, 13, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    doc.Range(0, 0).InsertParagraphBefore();
                    Word.Range range = doc.Paragraphs[1].Range.Duplicate;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    doc.Fields.Add(range, Word.WdFieldType.wdFieldTOC, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                });
        }

        private static bool ReplaceText(WdSession s, string path, string from, string to)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => ReplaceLate(doc, from, to))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, from);
                    if (para == null)
                    {
                        return false;
                    }

                    para.Range.Text = to + "\r";
                    return true;
                });
        }

        private static bool StyleHas(WdSession s, string path, string needle, string[] names)
        {
                string found = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    found = ParaStyleLate(doc, needle);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, needle);
                    if (para == null)
                    {
                        return false;
                    }

                    found = ParagraphStyleName(doc, para);
                    return found.Length > 0;
                });
            return ok && HasName(found, names);
        }

        private static bool StyleIs(WdSession s, string path, string needle, string[] names)
        {
            string found = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    found = ParaStyleLate(doc, needle);
                    return found.Length > 0;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, needle);
                    if (para == null)
                    {
                        return false;
                    }

                    found = ParagraphStyleName(doc, para);
                    return found.Length > 0;
                });
            return ok && NameIs(found, names);
        }

        private static string ParagraphStyleName(Word.Document doc, Word.Paragraph para)
        {
            object style = para.get_Style();
            Word.Style typed = style as Word.Style;
            if (typed == null && style != null)
            {
                int code;
                if (int.TryParse(Convert.ToString(style, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out code))
                {
                    try
                    {
                        typed = doc.Styles[(Word.WdBuiltinStyle)code];
                    }
                    catch (Exception)
                    {
                        typed = null;
                    }
                }
            }

            return typed == null ? "" : (typed.NameLocal ?? "");
        }

        private static bool SliceHas(WdSession s, string path, string paraNeedle, string slice, string[] names)
        {
            string found = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = SliceLate(doc, paraNeedle, slice);
                    found = StyleNameLate(doc, range);
                    return range != null;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    found = StyleName(range);
                    return true;
                });
            return ok && HasName(found, names);
        }

        private static bool SliceNthHas(WdSession s, string path, string paraNeedle, string slice, int nth, string[] names)
        {
            string found = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = NthLate(doc, paraNeedle, slice, nth);
                    found = StyleNameLate(doc, range);
                    return range != null;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Nth(doc, paraNeedle, slice, nth);
                    if (range == null)
                    {
                        return false;
                    }

                    found = StyleName(range);
                    return true;
                });
            return ok && HasName(found, names);
        }

        private static bool Bold(WdSession s, string path, string paraNeedle, string slice)
        {
            int bold = 0;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object range = SliceLate(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    bold = Convert.ToInt32(WdSession.ComGet(WdSession.ComGet(range, "Font"), "Bold"));
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    bold = range.Font.Bold;
                    return true;
                });
            }

            return bold == 1 || bold == -1;
        }

        private static float ReadSize(WdSession s, string path, string paraNeedle, string slice)
        {
            float size = 0;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object range = SliceLate(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    size = Convert.ToSingle(WdSession.ComGet(WdSession.ComGet(range, "Font"), "Size"), CultureInfo.InvariantCulture);
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    Word.Range range = Slice(doc, paraNeedle, slice);
                    if (range == null)
                    {
                        return false;
                    }

                    size = range.Font.Size;
                    return true;
                });
            }

            return size;
        }

        private static string FullXml(WdSession s, string path)
        {
            string xml = "";
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    xml = Convert.ToString(WdSession.ComGet(doc, "WordOpenXML")) ?? "";
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    xml = doc.WordOpenXML ?? "";
                    return true;
                });
            }

            return xml;
        }

        /// <summary>整篇没变，或只补进原先没有的内置样式。正文不变。原有样式定义不变，样式编号可以重排。可以因此多出 numbering.xml 和指向它的关系。</summary>
        private static bool OnlyBuiltinStyleFill(string before, string after)
        {
            if (before == after)
            {
                return true;
            }

            Dictionary<string, XElement> left = PackageRoots(before);
            Dictionary<string, XElement> right = PackageRoots(after);
            if (left.Count == 0)
            {
                return false;
            }

            foreach (string name in left.Keys)
            {
                if (!right.ContainsKey(name))
                {
                    return false;
                }
            }

            foreach (string name in right.Keys)
            {
                if (!left.ContainsKey(name) && !name.EndsWith("/numbering.xml", StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }
            }

            foreach (KeyValuePair<string, XElement> pair in left)
            {
                XElement other = right[pair.Key];
                if (pair.Key.EndsWith("styles.xml", StringComparison.OrdinalIgnoreCase))
                {
                    if (!StylesOnlyGrew(pair.Value, other))
                    {
                        return false;
                    }
                }
                else if (pair.Key.EndsWith("document.xml.rels", StringComparison.OrdinalIgnoreCase))
                {
                    if (!RelsOnlyAddedNumbering(pair.Value, other))
                    {
                        return false;
                    }
                }
                else if (!XNode.DeepEquals(pair.Value, other))
                {
                    return false;
                }
            }

            return true;
        }

        private static Dictionary<string, XElement> PackageRoots(string xml)
        {
            var map = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(xml))
            {
                return map;
            }

            XDocument doc;
            try
            {
                doc = XDocument.Parse(xml);
            }
            catch (Exception)
            {
                return map;
            }

            foreach (XElement part in doc.Descendants())
            {
                if (part.Name.LocalName != "part")
                {
                    continue;
                }

                string name = null;
                foreach (XAttribute attr in part.Attributes())
                {
                    if (attr.Name.LocalName == "name")
                    {
                        name = attr.Value;
                        break;
                    }
                }

                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                XElement root = part;
                foreach (XElement child in part.Elements())
                {
                    if (child.Name.LocalName != "xmlData")
                    {
                        continue;
                    }

                    foreach (XElement data in child.Elements())
                    {
                        root = data;
                        break;
                    }

                    break;
                }

                map[name] = root;
            }

            return map;
        }

        private static bool StylesOnlyGrew(XElement before, XElement after)
        {
            var beforeStyles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            var afterStyles = new Dictionary<string, XElement>(StringComparer.OrdinalIgnoreCase);
            if (!CollectStyles(before, beforeStyles) || !CollectStyles(after, afterStyles))
            {
                return false;
            }

            if (!SameNonStyleChildren(before, after))
            {
                return false;
            }

            foreach (KeyValuePair<string, XElement> pair in beforeStyles)
            {
                XElement other;
                if (!afterStyles.TryGetValue(pair.Key, out other) || !SameStyle(pair.Value, other))
                {
                    return false;
                }
            }

            return afterStyles.Count >= beforeStyles.Count;
        }

        private static bool CollectStyles(XElement styles, Dictionary<string, XElement> byName)
        {
            foreach (XElement child in styles.Elements())
            {
                if (child.Name.LocalName != "style")
                {
                    continue;
                }

                string name = null;
                foreach (XElement node in child.Elements())
                {
                    if (node.Name.LocalName != "name")
                    {
                        continue;
                    }

                    foreach (XAttribute attr in node.Attributes())
                    {
                        if (attr.Name.LocalName == "val")
                        {
                            name = attr.Value;
                            break;
                        }
                    }

                    break;
                }

                if (string.IsNullOrEmpty(name) || byName.ContainsKey(name))
                {
                    return false;
                }

                byName[name] = child;
            }

            return true;
        }

        private static bool SameStyle(XElement before, XElement after)
        {
            return XNode.DeepEquals(WithoutStyleId(before), WithoutStyleId(after));
        }

        private static XElement WithoutStyleId(XElement style)
        {
            var copy = new XElement(style);
            var drop = new List<XAttribute>();
            foreach (XAttribute attr in copy.Attributes())
            {
                if (attr.Name.LocalName == "styleId")
                {
                    drop.Add(attr);
                }
            }

            foreach (XAttribute attr in drop)
            {
                attr.Remove();
            }

            return copy;
        }

        private static bool RelsOnlyAddedNumbering(XElement before, XElement after)
        {
            var left = new Dictionary<string, XElement>(StringComparer.Ordinal);
            var right = new Dictionary<string, XElement>(StringComparer.Ordinal);
            if (!CollectRels(before, left) || !CollectRels(after, right))
            {
                return false;
            }

            foreach (string key in left.Keys)
            {
                if (!right.ContainsKey(key))
                {
                    return false;
                }
            }

            foreach (KeyValuePair<string, XElement> pair in right)
            {
                if (left.ContainsKey(pair.Key))
                {
                    continue;
                }

                string target = "";
                foreach (XAttribute attr in pair.Value.Attributes())
                {
                    if (attr.Name.LocalName == "Target")
                    {
                        target = attr.Value ?? "";
                        break;
                    }
                }

                if (target.IndexOf("numbering.xml", StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool CollectRels(XElement root, Dictionary<string, XElement> byId)
        {
            foreach (XElement child in root.Descendants())
            {
                if (child.Name.LocalName != "Relationship")
                {
                    continue;
                }

                string type = "";
                string target = "";
                foreach (XAttribute attr in child.Attributes())
                {
                    if (attr.Name.LocalName == "Type")
                    {
                        type = attr.Value ?? "";
                    }
                    else if (attr.Name.LocalName == "Target")
                    {
                        target = attr.Value ?? "";
                    }
                }

                string id = type + "|" + target;
                if (id == "|" || byId.ContainsKey(id))
                {
                    return false;
                }

                byId[id] = child;
            }

            return true;
        }

        private static bool SameNonStyleChildren(XElement before, XElement after)
        {
            var left = new List<XElement>();
            var right = new List<XElement>();
            foreach (XElement child in before.Elements())
            {
                if (child.Name.LocalName != "style")
                {
                    left.Add(child);
                }
            }

            foreach (XElement child in after.Elements())
            {
                if (child.Name.LocalName != "style")
                {
                    right.Add(child);
                }
            }

            if (left.Count != right.Count)
            {
                return false;
            }

            for (int i = 0; i < left.Count; i++)
            {
                if (!XNode.DeepEquals(left[i], right[i]))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool SetStyle(Word.Range range, string[] names)
        {
            foreach (string name in names)
            {
                try
                {
                    object styleRef = name;
                    range.set_Style(ref styleRef);
                    return true;
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static Word.Paragraph FindPara(Word.Document doc, string needle)
        {
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                string text = para.Range.Text ?? "";
                if (text.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return para;
                }
            }

            return null;
        }

        private static Word.Range Slice(Word.Document doc, string paraNeedle, string slice)
        {
            Word.Paragraph para = FindPara(doc, paraNeedle);
            if (para == null)
            {
                return null;
            }

            string text = para.Range.Text ?? "";
            int at = text.IndexOf(slice, StringComparison.Ordinal);
            if (at < 0)
            {
                return null;
            }

            return doc.Range(para.Range.Start + at, para.Range.Start + at + slice.Length);
        }

        private static Word.Range Nth(Word.Document doc, string paraNeedle, string slice, int nth)
        {
            Word.Paragraph para = FindPara(doc, paraNeedle);
            if (para == null)
            {
                return null;
            }

            string text = para.Range.Text ?? "";
            int seen = 0;
            int at = 0;
            while (at <= text.Length - slice.Length)
            {
                int found = text.IndexOf(slice, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    return null;
                }

                seen++;
                if (seen == nth)
                {
                    return doc.Range(para.Range.Start + found, para.Range.Start + found + slice.Length);
                }

                at = found + slice.Length;
            }

            return null;
        }

        private static string StyleName(Word.Range range)
        {
            try
            {
                object style = range.get_Style();
                Word.Style typed = style as Word.Style;
                if (typed != null)
                {
                    return typed.NameLocal ?? "";
                }

                return Convert.ToString(style) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool PaintParagraphLate(object doc, string needle, string[] names)
        {
            object para = FindParaLate(doc, needle);
            if (para == null)
            {
                return false;
            }

            return SetStyleLate(WdSession.ComGet(para, "Range"), names);
        }

        private static bool PaintSliceLate(object doc, string paraNeedle, string slice, string[] names)
        {
            object range = SliceLate(doc, paraNeedle, slice);
            return range != null && SetStyleLate(range, names);
        }

        private static bool PaintNthLate(object doc, string paraNeedle, string slice, int nth, string[] names)
        {
            object range = NthLate(doc, paraNeedle, slice, nth);
            return range != null && SetStyleLate(range, names);
        }

        private static bool PaintSpaceLate(object doc, string needle, float before)
        {
            object para = FindParaLate(doc, needle);
            if (para == null)
            {
                return false;
            }

            WdSession.ComSet(WdSession.ComGet(para, "Format"), "SpaceBefore", before);
            return true;
        }

        private static bool ReplaceLate(object doc, string from, string to)
        {
            object para = FindParaLate(doc, from);
            if (para == null)
            {
                return false;
            }

            WdSession.ComSet(WdSession.ComGet(para, "Range"), "Text", to + "\r");
            return true;
        }

        private static string ParaStyleLate(object doc, string needle)
        {
            object para = FindParaLate(doc, needle);
            if (para == null)
            {
                return "";
            }

            object style = WdSession.ComGet(para, "Style");
            string name = Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
            if (name.Length > 0)
            {
                return name;
            }

            int code;
            if (!int.TryParse(Convert.ToString(style, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out code))
            {
                return "";
            }

            try
            {
                object styles = WdSession.ComGet(doc, "Styles");
                object resolved = WdSession.ComCall(styles, "Item", code);
                return Convert.ToString(WdSession.ComGet(resolved, "NameLocal")) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static object FindParaLate(object doc, string needle)
        {
            object paras = WdSession.ComGet(doc, "Paragraphs");
            int count = Convert.ToInt32(WdSession.ComGet(paras, "Count"));
            for (int i = 1; i <= count; i++)
            {
                object para = WdSession.ComCall(paras, "Item", i);
                string text = Convert.ToString(WdSession.ComGet(WdSession.ComGet(para, "Range"), "Text")) ?? "";
                if (text.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return para;
                }
            }

            return null;
        }

        private static object SliceLate(object doc, string paraNeedle, string slice)
        {
            object para = FindParaLate(doc, paraNeedle);
            if (para == null)
            {
                return null;
            }

            object range = WdSession.ComGet(para, "Range");
            string text = Convert.ToString(WdSession.ComGet(range, "Text")) ?? "";
            int at = text.IndexOf(slice, StringComparison.Ordinal);
            if (at < 0)
            {
                return null;
            }

            int start = Convert.ToInt32(WdSession.ComGet(range, "Start"));
            return WdSession.ComCall(doc, "Range", start + at, start + at + slice.Length);
        }

        private static object NthLate(object doc, string paraNeedle, string slice, int nth)
        {
            object para = FindParaLate(doc, paraNeedle);
            if (para == null)
            {
                return null;
            }

            object range = WdSession.ComGet(para, "Range");
            string text = Convert.ToString(WdSession.ComGet(range, "Text")) ?? "";
            int start = Convert.ToInt32(WdSession.ComGet(range, "Start"));
            int seen = 0;
            int at = 0;
            while (at <= text.Length - slice.Length)
            {
                int found = text.IndexOf(slice, at, StringComparison.Ordinal);
                if (found < 0)
                {
                    return null;
                }

                seen++;
                if (seen == nth)
                {
                    return WdSession.ComCall(doc, "Range", start + found, start + found + slice.Length);
                }

                at = found + slice.Length;
            }

            return null;
        }

        private static bool SetStyleLate(object range, string[] names)
        {
            foreach (string name in names)
            {
                try
                {
                    WdSession.ComSet(range, "Style", name);
                    return true;
                }
                catch (Exception)
                {
                }
            }

            return false;
        }

        private static string StyleNameLate(object doc, object range)
        {
            if (range == null)
            {
                return "";
            }

            object style;
            try
            {
                style = WdSession.ComGet(range, "Style");
            }
            catch (Exception)
            {
                return "";
            }

            if (style == null)
            {
                return "";
            }

            try
            {
                string name = Convert.ToString(WdSession.ComGet(style, "NameLocal")) ?? "";
                if (name.Length > 0)
                {
                    return name;
                }
            }
            catch (Exception)
            {
            }

            int code;
            if (!int.TryParse(Convert.ToString(style, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out code))
            {
                return "";
            }

            try
            {
                object styles = WdSession.ComGet(doc, "Styles");
                object resolved = WdSession.ComCall(styles, "Item", code);
                return Convert.ToString(WdSession.ComGet(resolved, "NameLocal")) ?? "";
            }
            catch (Exception)
            {
                return "";
            }
        }

        private static bool HasName(string found, string[] names)
        {
            if (string.IsNullOrEmpty(found))
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

        private static bool NameIs(string found, string[] names)
        {
            if (string.IsNullOrEmpty(found))
            {
                return false;
            }

            foreach (string name in names)
            {
                if (string.Equals(found.Trim(), name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TocOk(JArray styles)
        {
            foreach (JObject row in Each(styles))
            {
                string structure = Str(row, "structure");
                if (structure.IndexOf("TOC", StringComparison.OrdinalIgnoreCase) < 0
                    && structure.IndexOf("目录", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                if (Eq(row, "support", "paragraph") && !structure.StartsWith("heading_", StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool ApplyMatches(JArray styles)
        {
            foreach (JObject row in Each(styles))
            {
                if (Str(row, "apply") != Str(row, "support"))
                {
                    return false;
                }
            }

            return styles.Count > 0;
        }

        private static Dictionary<string, object> Map(params object[] pairs)
        {
            var map = new Dictionary<string, object>();
            for (int i = 0; i + 1 < pairs.Length; i += 2)
            {
                map[(string)pairs[i]] = pairs[i + 1];
            }

            return map;
        }

        private static JObject Parse(ToolResult result)
        {
            if (result == null || result.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(result.Data));
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static JArray Arr(JObject data, string key)
        {
            return data == null ? new JArray() : (data[key] as JArray ?? new JArray());
        }

        private static JObject Obj(JObject data, string key)
        {
            return data == null ? null : data[key] as JObject;
        }

        private static JObject Find(JArray rows, string structure)
        {
            foreach (JObject row in Each(rows))
            {
                if (string.Equals(Str(row, "structure"), structure, StringComparison.OrdinalIgnoreCase))
                {
                    return row;
                }
            }

            return null;
        }

        private static JObject SpanText(JArray spans, string text)
        {
            foreach (JObject span in Each(spans))
            {
                if (TextValue(span) == text || (TextValue(span) != null && TextValue(span).IndexOf(text, StringComparison.Ordinal) >= 0 && text.Length >= 2))
                {
                    return span;
                }
            }

            return null;
        }

        private static string TextValue(JObject span)
        {
            if (span == null || span["text"] == null)
            {
                return null;
            }

            JToken text = span["text"];
            if (text.Type == JTokenType.Object)
            {
                return text["value"] == null ? null : text["value"].ToString();
            }

            return text.ToString();
        }

        private static int TextIndex(JObject span)
        {
            if (span == null || !(span["text"] is JObject text) || text["index"] == null)
            {
                return 0;
            }

            return text["index"].Value<int>();
        }

        private static JObject SpanByCode(JArray spans, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return null;
            }

            foreach (JObject span in Each(spans))
            {
                if (Str(span, "target_codes") == code)
                {
                    return span;
                }
            }

            return null;
        }

        private static JObject First(JArray rows)
        {
            return rows.Count == 0 ? null : rows[0] as JObject;
        }

        private static IEnumerable<JObject> Each(JArray rows)
        {
            if (rows == null)
            {
                yield break;
            }

            foreach (JToken token in rows)
            {
                if (token is JObject row)
                {
                    yield return row;
                }
            }
        }

        private static bool OkFlag(JObject data)
        {
            if (data == null)
            {
                return false;
            }

            if (data["ok"] != null && data["ok"].Type == JTokenType.Boolean)
            {
                return data["ok"].Value<bool>();
            }

            return data["_success"] != null && data["_success"].Type == JTokenType.Boolean && data["_success"].Value<bool>();
        }

        private static bool Eq(JObject data, string key, string expected)
        {
            return data != null && string.Equals(Str(data, key), expected, StringComparison.OrdinalIgnoreCase);
        }

        private static string Str(JObject data, string key)
        {
            if (data == null || data[key] == null)
            {
                return "";
            }

            return data[key].ToString();
        }

        private static bool Flag(JObject data, string key)
        {
            return data != null && data[key] != null && data[key].Type == JTokenType.Boolean && data[key].Value<bool>();
        }

        private static bool Near(JObject data, string key, double expected, double tol)
        {
            if (data == null || data[key] == null)
            {
                return false;
            }

            double value;
            return double.TryParse(data[key].ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                && Math.Abs(value - expected) <= tol;
        }

        private static bool Blob(JObject data, string needle)
        {
            return data != null && Dump(data).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string Dump(JToken token)
        {
            if (token == null)
            {
                return "(空)";
            }

            string text = token.ToString(Formatting.None);
            return text.Length <= 500 ? text : text.Substring(0, 500);
        }

        private static string Trim(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Length <= 180 ? text : text.Substring(0, 180);
        }

        private static bool HostHasStyle(WdSession s, string path, params string[] names)
        {
            bool found = false;
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    object styles = WdSession.ComGet(doc, "Styles");
                    foreach (string name in names)
                    {
                        try
                        {
                            object style = WdSession.ComCall(styles, "Item", name);
                            if (style != null)
                            {
                                found = true;
                                return true;
                            }
                        }
                        catch (Exception)
                        {
                        }
                    }

                    return false;
                });
                return found;
            }

            s.OnWordDoc(path, doc =>
            {
                foreach (string name in names)
                {
                    try
                    {
                        Word.Style style = doc.Styles[name];
                        if (style != null)
                        {
                            found = true;
                            return true;
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return false;
            });
            return found;
        }

        private static void Ok(TestRun run, string id, bool pass, string detail)
        {
            if (Only != null && !Only.Contains(id))
            {
                return;
            }

            if (pass)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, detail ?? "");
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                if (Only != null && !Only.Contains(id))
                {
                    continue;
                }

                run.CaseFail(id, reason);
            }
        }
    }
}
