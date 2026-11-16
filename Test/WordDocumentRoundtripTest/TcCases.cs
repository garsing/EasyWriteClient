using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>目录域和目录内容控件。对照 docs/测试/word测试/目录。</summary>
    internal static class TcCases
    {
        public const int Target = 90;

        private static readonly Regex Token = new Regex(
            @"\[目录域 (ewC_[A-Za-z0-9]{8,32})\]",
            RegexOptions.Compiled);

        private static readonly string[] HeadingNames = { "标题 1", "heading 1", "Heading 1" };

        private const string GalleryXml =
            "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:sdt><w:sdtPr><w:id w:val=\"1051109989\"/><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:t>目录</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 1 级)</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 2 级)</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 3 级)</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 1 级)</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 2 级)</w:t></w:r></w:p><w:p><w:r><w:t>键入章标题(第 3 级)</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>后文你好啊。</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";

        private static HashSet<string> Only;

        public static void Run(TestRun run, string host, string filter)
        {
            Only = null;
            if (!string.IsNullOrWhiteSpace(filter)
                && !filter.Equals("tc", StringComparison.OrdinalIgnoreCase)
                && !filter.Equals("tc-", StringComparison.OrdinalIgnoreCase))
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
            RunSplit(run);
            RunConsolidate(run);
            RunStandard(run, s);
            RunNotoc(run, s);
            RunTwo(run, s);
            RunCell(run, s);
            RunNest(run, s);
            RunOrder(run, s);
            RunPeriod(run, s);
            RunNotTocField(run, s);
            RunBothSides(run, s);
            RunCrNext(run, s);
            RunCrMissing(run, s);
            RunCrBack(run, s);
            RunPlain(run, s);
            RunControl(run, s);
            RunControlField(run, s);
        }

        private static void RunSplit(TestRun run)
        {
            if (!Want("tc-para-empty-keep", "tc-para-tail-drop", "tc-para-blank", "tc-para-mid"))
            {
                return;
            }

            string marker = TocFieldStamp.FormatMarker("ewC_abcdefgh");
            List<ParagraphBlock> empty = ParasThroughSeg(marker + "\r");
            Ok(run, "tc-para-empty-keep",
                empty.Count == 1 && empty[0].Text.Length == 0 && empty[0].TrailingCr,
                "n=" + empty.Count);
            List<ParagraphBlock> tail = ParasThroughSeg("甲\r" + marker);
            Ok(run, "tc-para-tail-drop",
                tail.Count == 1 && tail[0].Text == "甲",
                "n=" + tail.Count + " t=" + (tail.Count > 0 ? tail[0].Text : ""));
            List<ParagraphBlock> blank = SplitTextByParagraphs.Split("\r");
            Ok(run, "tc-para-blank", blank.Count == 1 && blank[0].Text.Length == 0, "n=" + blank.Count);
            List<ParagraphBlock> mid = ParasThroughSeg("甲" + marker + "乙\r");
            Ok(run, "tc-para-mid",
                mid.Count == 2 && mid[0].Text == "甲" && mid[1].Text == "乙",
                "n=" + mid.Count + " t=" + string.Join("|", mid.ConvertAll(block => block.Text)));
        }

        private static List<ParagraphBlock> ParasThroughSeg(string text)
        {
            var blocks = new List<ParagraphBlock>();
            List<string> parts = WordDocumentExtractor.SplitTextByCellsForMapping(text, out List<int> _);
            for (int i = 1; i < parts.Count; i += 2)
            {
                blocks.AddRange(SplitTextByParagraphs.Split(parts[i]));
            }

            return blocks;
        }

        private static void RunConsolidate(TestRun run)
        {
            if (!Want(
                "tc-seg-three", "tc-seg-head", "tc-seg-tail", "tc-seg-cr",
                "tc-seg-two", "tc-seg-same", "tc-seg-image", "tc-seg-cell", "tc-seg-pair", "tc-seg-nest",
                "tc-col-one", "tc-col-two", "tc-col-empty", "tc-col-keep", "tc-col-story", "tc-col-agent"))
            {
                return;
            }

            string a = TocFieldStamp.FormatMarker("ewC_aaaaaaaa");
            string b = TocFieldStamp.FormatMarker("ewC_bbbbbbbb");
            string one = "前文\r" + a + "目录行\r" + a + "后文。";
            List<string> three = SegTexts(one);
            Ok(run, "tc-seg-three",
                three.Count == 3 && three[0] == "前文\r" && three[1] == "目录行\r" && three[2] == "后文。",
                JoinSegs(three));

            string head = a + "目录\r" + a + "后";
            List<string> headSegs = SegTexts(head);
            Ok(run, "tc-seg-head",
                headSegs.Count == 2 && headSegs[0] == "目录\r" && headSegs[1] == "后"
                && GapTexts(head)[0].StartsWith(a, StringComparison.Ordinal),
                JoinSegs(headSegs));

            string tail = "甲\r" + a;
            List<string> tailSegs = SegTexts(tail);
            List<ParagraphBlock> tailParas = ParasThroughSeg(tail);
            Ok(run, "tc-seg-tail",
                tailSegs.Count == 1 && tailSegs[0] == "甲\r"
                && tailParas.Count == 1 && tailParas[0].Text == "甲",
                JoinSegs(tailSegs));

            string cr = "甲" + a + "\r乙" + a + "后";
            List<string> crSegs = SegTexts(cr);
            List<string> crGaps = GapTexts(cr);
            string crGap = crGaps.Count > 1 ? crGaps[1] : "";
            Ok(run, "tc-seg-cr",
                crSegs.Count == 3 && crSegs[0] == "甲" && crSegs[1] == "\r乙" && crGap == a,
                "gap=" + crGap + " segs=" + JoinSegs(crSegs));

            string two = "前文\r" + a + "目录甲\r" + a + "中\r" + b + "目录乙\r" + b + "后文。";
            List<string> twoSegs = SegTexts(two);
            Ok(run, "tc-seg-two",
                twoSegs.Count == 5
                && twoSegs[0] == "前文\r" && twoSegs[1] == "目录甲\r" && twoSegs[2] == "中\r"
                && twoSegs[3] == "目录乙\r" && twoSegs[4] == "后文。",
                JoinSegs(twoSegs));

            string same = "前文\r" + a + "目录行\r" + a + "中\r" + b + "目录行\r" + b + "后文。";
            List<string> sameSegs = SegTexts(same);
            Ok(run, "tc-seg-same",
                sameSegs.Count == 5 && sameSegs[1] == "目录行\r" && sameSegs[3] == "目录行\r" && sameSegs[2] == "中\r",
                JoinSegs(sameSegs));

            string image = "<image id=\"i1\" src=\"a\"/>\r" + a + "目录\r" + a + "后";
            List<string> imageSegs = SegTexts(image);
            Ok(run, "tc-seg-image",
                imageSegs.Count == 2 && imageSegs[0] == "目录\r" && imageSegs[1] == "后",
                JoinSegs(imageSegs));

            string cell = "前<table id=\"t\"><td>格内" + a + "目录\r" + a + "格尾</td></table>后";
            List<string> cellSegs = SegTexts(cell);
            Ok(run, "tc-seg-cell",
                cellSegs.Count == 5
                && cellSegs[0] == "前" && cellSegs[1] == "格内" && cellSegs[2] == "目录\r"
                && cellSegs[3] == "格尾" && cellSegs[4] == "后",
                JoinSegs(cellSegs));

            List<string> pairIds = MarkerIds(string.Join("", GapTexts(same)));
            bool markerInSeg = false;
            foreach (string seg in sameSegs)
            {
                if ((seg ?? "").IndexOf("<toc", StringComparison.Ordinal) >= 0)
                {
                    markerInSeg = true;
                }
            }

            Ok(run, "tc-seg-pair",
                !markerInSeg && pairIds.Count == 4
                && pairIds[0] == "ewC_aaaaaaaa" && pairIds[1] == "ewC_aaaaaaaa"
                && pairIds[2] == "ewC_bbbbbbbb" && pairIds[3] == "ewC_bbbbbbbb",
                string.Join(",", pairIds));

            string nest = "外前<table id=\"o\"><td>外格<table id=\"i\"><td>内" + a + "目录" + a + "内尾</td></table>外尾</td></table>外后";
            List<string> nestSegs = SegTexts(nest);
            int nestAt = nestSegs.IndexOf("目录");
            Ok(run, "tc-seg-nest",
                nestAt > 0 && nestAt + 1 < nestSegs.Count
                && nestSegs[nestAt - 1] == "内" && nestSegs[nestAt + 1] == "内尾"
                && nestSegs.Contains("外后"),
                JoinSegs(nestSegs));

            string wrapped = "前文" + a + "[S_abcde-start]目录行[S_abcde-end]" + a + "后文";
            Ok(run, "tc-col-one",
                TocFieldStamp.CollapseMarked(wrapped) == "前文[目录域 ewC_aaaaaaaa]后文",
                TocFieldStamp.CollapseMarked(wrapped));

            string twoWrap = "前文" + a + "目录甲" + a + "中" + b + "目录乙" + b + "后文";
            Ok(run, "tc-col-two",
                TocFieldStamp.CollapseMarked(twoWrap) == "前文[目录域 ewC_aaaaaaaa]中[目录域 ewC_bbbbbbbb]后文",
                TocFieldStamp.CollapseMarked(twoWrap));

            Ok(run, "tc-col-empty",
                TocFieldStamp.CollapseMarked(a + a) == "[目录域 ewC_aaaaaaaa]",
                TocFieldStamp.CollapseMarked(a + a));

            string lone = "前文" + a + "后文";
            Ok(run, "tc-col-keep",
                TocFieldStamp.CollapseMarked(lone) == lone,
                TocFieldStamp.CollapseMarked(lone));

            string storyRaw = "前文\r" + a + "目录行\t1\r" + a + "你好啊。";
            string story = DisplayStoryConvert.ToStory(storyRaw);
            Ok(run, "tc-col-story",
                story.IndexOf("目录行", StringComparison.Ordinal) >= 0
                && story.IndexOf("你好啊", StringComparison.Ordinal) >= 0
                && story.IndexOf("ewC_", StringComparison.Ordinal) < 0
                && story.IndexOf("<toc", StringComparison.Ordinal) < 0,
                story);

            string agentStory = DisplayStoryConvert.ToStory(TocFieldStamp.CollapseMarked(storyRaw));
            Ok(run, "tc-col-agent",
                agentStory.IndexOf("你好啊", StringComparison.Ordinal) >= 0
                && agentStory.IndexOf("目录行", StringComparison.Ordinal) < 0
                && agentStory.IndexOf("ewC_", StringComparison.Ordinal) < 0,
                agentStory);
        }

        private static List<string> SegTexts(string text)
        {
            var segs = new List<string>();
            List<string> parts = WordDocumentExtractor.SplitTextByCellsForMapping(text, out List<int> _);
            for (int i = 1; i < parts.Count; i += 2)
            {
                segs.Add(parts[i]);
            }

            return segs;
        }

        private static List<string> GapTexts(string text)
        {
            var gaps = new List<string>();
            List<string> parts = WordDocumentExtractor.SplitTextByCellsForMapping(text, out List<int> _);
            for (int i = 0; i < parts.Count; i += 2)
            {
                gaps.Add(parts[i]);
            }

            return gaps;
        }

        private static List<string> MarkerIds(string text)
        {
            var ids = new List<string>();
            foreach (Match match in TocFieldStamp.Marker.Matches(text ?? ""))
            {
                ids.Add(match.Groups[1].Value);
            }

            return ids;
        }

        private static string JoinSegs(List<string> segs)
        {
            return segs == null ? "" : string.Join(" | ", segs.ConvertAll(seg => (seg ?? "").Replace("\r", "\\r")));
        }

        private static string ParaHolding(string display, string text)
        {
            foreach (Match match in Regex.Matches(display ?? "", @"\[(P_[0-9A-Za-z]{5}(?:_\d+)?) start\]([\s\S]*?)\[\1 end\]"))
            {
                if ((match.Groups[2].Value ?? "").IndexOf(text, StringComparison.Ordinal) >= 0)
                {
                    return match.Groups[2].Value ?? "";
                }
            }

            return "";
        }

        private static void RunStandard(TestRun run, WdSession s)
        {
            if (!Want(
                "tc-id-stamp", "tc-id-reuse", "tc-id-shape",
                "tc-sent-hidden", "tc-sent-no-lone", "tc-sent-body", "tc-sent-not-glued", "tc-sent-hyperlink",
                "tc-para-order", "tc-para-stable", "tc-cr-has",
                "tc-fold-sent", "tc-fold-para", "tc-fold-chunk", "tc-fold-lines", "tc-fold-body",
                "tc-fold-full", "tc-fold-story", "tc-fold-inside",
                "tc-op-sent", "tc-op-para", "tc-op-style", "tc-op-ewc", "tc-op-body",
                "tc-span-hide", "tc-span-style", "tc-nth",
                "tc-mark-pair", "tc-mark-between", "tc-glue-store", "tc-para-store"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r你好啊。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !UpdateFields(s, path))
            {
                FailMany(run, "铺标准目录稿失败",
                    "tc-id-stamp", "tc-id-reuse", "tc-id-shape",
                    "tc-sent-hidden", "tc-sent-no-lone", "tc-sent-body", "tc-sent-not-glued", "tc-sent-hyperlink",
                    "tc-para-order", "tc-para-stable", "tc-cr-has",
                    "tc-fold-sent", "tc-fold-para", "tc-fold-chunk", "tc-fold-lines", "tc-fold-body",
                    "tc-fold-full", "tc-fold-story", "tc-fold-inside",
                    "tc-op-sent", "tc-op-para", "tc-op-style", "tc-op-ewc", "tc-op-body",
                    "tc-span-hide", "tc-span-style", "tc-nth",
                    "tc-mark-pair", "tc-mark-between", "tc-glue-store", "tc-para-store");
                return;
            }

            try
            {
                int before = ParagraphCount(s, path);
                string sent = Sentence(s, id);
                string para = Paragraph(s, id);
                Match token = Token.Match(sent ?? "");
                string name = token.Success ? token.Groups[1].Value : "";
                string bookmark = BookmarkAtField(s, path);
                Ok(run, "tc-id-stamp",
                    token.Success && name == bookmark && bookmark.StartsWith("ewC_", StringComparison.Ordinal),
                    "sent=" + sent + " mark=" + bookmark);
                string sent2 = Sentence(s, id);
                Match again = Token.Match(sent2 ?? "");
                Ok(run, "tc-id-reuse",
                    again.Success && again.Groups[1].Value == name && CountBookmarks(s, path, name) == 1,
                    "again=" + (again.Success ? again.Groups[1].Value : ""));
                Ok(run, "tc-id-shape", Token.IsMatch(name) || Token.IsMatch("[目录域 " + name + "]"), "name=" + name);
                Ok(run, "tc-sent-hidden",
                    (sent ?? "").IndexOf("<toc", StringComparison.Ordinal) < 0
                    && (para ?? "").IndexOf("<toc", StringComparison.Ordinal) < 0,
                    "sent=" + sent);
                Ok(run, "tc-sent-no-lone",
                    token.Success && (sent ?? "").IndexOf("歧义无法操作", StringComparison.Ordinal) < 0,
                    sent);
                string bodyCode = Sent(sent, "你好啊");
                Ok(run, "tc-sent-body",
                    bodyCode.StartsWith("S_", StringComparison.Ordinal) && !InsideToken(sent, "你好啊"),
                    "code=" + bodyCode + " sent=" + sent);
                Ok(run, "tc-sent-not-glued",
                    bodyCode.StartsWith("S_", StringComparison.Ordinal) && !SentenceHolds(sent, "你好啊", "\t"),
                    sent);
                Ok(run, "tc-sent-hyperlink", Token.Matches(sent ?? "").Count == 1, sent);
                Ok(run, "tc-para-order", BodyParagraphAfterToc(s, path, "你好啊"), "count=" + ParagraphCount(s, path));
                int after = ParagraphCount(s, path);
                Ok(run, "tc-para-stable", before > 0 && before == after, before + "/" + after);
                Ok(run, "tc-cr-has", before > 0 && before == after, before + "/" + after);
                Ok(run, "tc-fold-sent",
                    Token.Matches(sent ?? "").Count == 1 && Outside(sent).IndexOf("\t", StringComparison.Ordinal) < 0,
                    sent);
                Ok(run, "tc-fold-para",
                    Token.Matches(para ?? "").Count == 1 && Outside(para).IndexOf("\t", StringComparison.Ordinal) < 0
                    && (para ?? "").IndexOf("你好啊", StringComparison.Ordinal) >= 0,
                    para);
                Ok(run, "tc-fold-lines", Token.Matches(para ?? "").Count == 1, para);
                Ok(run, "tc-fold-body",
                    (para ?? "").IndexOf("[P_", StringComparison.Ordinal) >= 0
                    && (para ?? "").IndexOf("你好啊", StringComparison.Ordinal) >= 0,
                    para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                Ok(run, "tc-fold-chunk",
                    Token.IsMatch(agent ?? "") && (agent ?? "").IndexOf("你好啊", StringComparison.Ordinal) >= 0
                    && Outside(agent).IndexOf("\t", StringComparison.Ordinal) < 0,
                    agent);
                Ok(run, "tc-fold-full",
                    (full ?? "").IndexOf("第一章", StringComparison.Ordinal) >= 0
                    && Outside(agent).IndexOf("\t", StringComparison.Ordinal) < 0,
                    "full=" + Brief(full) + " agent=" + Brief(agent));
                Ok(run, "tc-fold-story",
                    (agent ?? "").IndexOf("<toc", StringComparison.Ordinal) < 0
                    && (full ?? "").IndexOf("第一章", StringComparison.Ordinal) >= 0,
                    Brief(full));
                List<string> pair = MarkerIds(full);
                int openAt = (full ?? "").IndexOf("<toc", StringComparison.Ordinal);
                int closeAt = openAt < 0 ? -1 : (full ?? "").IndexOf("<toc", openAt + 4, StringComparison.Ordinal);
                string between = openAt >= 0 && closeAt > openAt
                    ? (full ?? "").Substring(openAt, closeAt - openAt)
                    : "";
                string afterMark = closeAt < 0 ? "" : (full ?? "").Substring(closeAt);
                Ok(run, "tc-mark-pair",
                    pair.Count == 2 && pair[0] == pair[1] && pair[0] == bookmark
                    && (agent ?? "").IndexOf("<toc", StringComparison.Ordinal) < 0,
                    string.Join(",", pair) + " mark=" + bookmark);
                Ok(run, "tc-mark-between",
                    between.IndexOf("第一章", StringComparison.Ordinal) >= 0
                    && afterMark.IndexOf("你好啊", StringComparison.Ordinal) >= 0
                    && between.IndexOf("你好啊", StringComparison.Ordinal) < 0,
                    Brief(full));
                bool glueOk = false;
                string glueDetail = "";
                foreach (IReadOnlyList<string> mapping in DocumentState.SentenceNameMapping)
                {
                    string content = mapping != null && mapping.Count >= 3 ? mapping[2] ?? "" : "";
                    if (content.IndexOf("你好啊", StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    glueOk = content.IndexOf('\t') < 0 && content.IndexOf("<toc", StringComparison.Ordinal) < 0;
                    glueDetail = content;
                    break;
                }

                Ok(run, "tc-glue-store", glueOk, glueDetail);
                string held = ParaHolding(full, "你好啊");
                Ok(run, "tc-para-store",
                    held.IndexOf("你好啊", StringComparison.Ordinal) >= 0
                    && held.IndexOf("第一章", StringComparison.Ordinal) < 0
                    && held.IndexOf("<toc", StringComparison.Ordinal) < 0,
                    held);
                string interior = FirstInteriorParagraph(s, path);
                Ok(run, "tc-fold-inside",
                    interior.StartsWith("P_", StringComparison.Ordinal) && (para ?? "").IndexOf(interior, StringComparison.Ordinal) < 0,
                    "para=" + interior);
                string sentErr = Replace(s, id, FirstInteriorSentence(s, path), "改掉");
                Ok(run, "tc-op-sent", sentErr.IndexOf("目录域里面的编码", StringComparison.Ordinal) >= 0, sentErr);
                string paraErr = Align(s, id, interior);
                Ok(run, "tc-op-para", paraErr.IndexOf("目录域里面的编码", StringComparison.Ordinal) >= 0, paraErr);
                string styleErr = StyleInterior(s, path);
                Ok(run, "tc-op-style", styleErr.IndexOf("目录域里面的编码", StringComparison.Ordinal) >= 0, styleErr);
                string ewcErr = Replace(s, id, name, "改掉");
                Ok(run, "tc-op-ewc",
                    ewcErr.Length > 0 && ewcErr.IndexOf("目录域里面的编码", StringComparison.Ordinal) < 0,
                    ewcErr);
                string freshBody = Sent(Sentence(s, id), "你好");
                string bodyErr = Replace(s, id, freshBody, "你好呀。");
                Ok(run, "tc-op-body", bodyErr.Length == 0, bodyErr);
                JArray spans = Spans(s, id);
                Ok(run, "tc-span-hide",
                    !SpanHas(spans, "\t") && SpanByCode(spans, Sent(Sentence(s, id), "你好")) != null,
                    Dump(spans));
                JArray styles = Styles(s, id);
                Ok(run, "tc-span-style",
                    FindStyle(styles, "目录 1") != null || FindStyle(styles, "TOC 1") != null,
                    Dump(styles));
                string visible = Outside(Paragraph(s, id));
                Ok(run, "tc-nth",
                    CountOf(visible, "第一章标题") == 1,
                    visible);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunNotoc(TestRun run, WdSession s)
        {
            if (!Want("tc-id-notoc"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r你好啊。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !UpdateFields(s, path) || !PlantTocAtStart(s, path))
            {
                Fail(run, "tc-id-notoc", "铺稿失败");
                return;
            }

            try
            {
                bool before = BookmarkExists(s, path, "_Toc245561084");
                Match token = Token.Match(Sentence(s, id) ?? "");
                string name = token.Success ? token.Groups[1].Value : "";
                bool after = BookmarkExists(s, path, "_Toc245561084");
                Ok(run, "tc-id-notoc",
                    token.Success && name.StartsWith("ewC_", StringComparison.Ordinal)
                    && after
                    && !string.Equals(name, "_Toc245561084", StringComparison.Ordinal),
                    name + " before=" + before + " after=" + after + " " + PlantNote + " marks=" + BookmarkNames(s, path));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunTwo(TestRun run, WdSession s)
        {
            if (!Want("tc-id-two", "tc-fold-two", "tc-fold-disagree", "tc-mark-order", "tc-mark-mid"))
            {
                return;
            }

            const string twoXml = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>目录甲</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p><w:p><w:r><w:t>中间正文。</w:t></w:r></w:p><w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>目录乙</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p><w:p><w:r><w:t>你好啊。</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
            string path, id;
            if (!Open(s, "占位。", out path, out id) || !InsertXml(s, path, twoXml))
            {
                FailMany(run, "铺两份目录失败", "tc-id-two", "tc-fold-two", "tc-fold-disagree", "tc-mark-order", "tc-mark-mid");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                MatchCollection marks = Token.Matches(para ?? "");
                bool distinct = marks.Count >= 2 && marks[0].Groups[1].Value != marks[1].Groups[1].Value;
                int first = (para ?? "").IndexOf("[目录域", StringComparison.Ordinal);
                int mid = (para ?? "").IndexOf("中间正文", StringComparison.Ordinal);
                int second = first < 0 ? -1 : (para ?? "").IndexOf("[目录域", first + 4, StringComparison.Ordinal);
                Ok(run, "tc-id-two", distinct && first >= 0 && second > first, para);
                Ok(run, "tc-fold-two", distinct && mid > first && (second < 0 || mid < second), para);
                Ok(run, "tc-fold-disagree",
                    mid >= 0 && (para ?? "").IndexOf("[P_", mid, StringComparison.Ordinal) >= 0,
                    para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                List<string> ids = MarkerIds(full);
                List<string> books = FieldBookmarks(s, path);
                MatchCollection markerHits = TocFieldStamp.Marker.Matches(full ?? "");
                string midSpan = "";
                if (markerHits.Count >= 3)
                {
                    int from = markerHits[1].Index + markerHits[1].Length;
                    int to = markerHits[2].Index;
                    if (to > from)
                    {
                        midSpan = (full ?? "").Substring(from, to - from);
                    }
                }

                Ok(run, "tc-mark-order",
                    ids.Count == 4 && ids[0] == ids[1] && ids[2] == ids[3] && ids[0] != ids[2]
                    && books.Count >= 2 && ids[0] == books[0] && ids[2] == books[1],
                    string.Join(",", ids) + " books=" + string.Join(",", books));
                Ok(run, "tc-mark-mid",
                    midSpan.IndexOf("中间正文", StringComparison.Ordinal) >= 0
                    && midSpan.IndexOf("目录甲", StringComparison.Ordinal) < 0
                    && midSpan.IndexOf("目录乙", StringComparison.Ordinal) < 0,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCell(TestRun run, WdSession s)
        {
            if (!Want("tc-id-cell", "tc-fold-table-mark", "tc-mark-cell"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r格外正文。", out path, out id) || !StyleFirst(s, path) || !AddTocInCell(s, path) || !UpdateFields(s, path))
            {
                FailMany(run, "铺单元格目录失败", "tc-id-cell", "tc-fold-table-mark", "tc-mark-cell");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                Match token = Token.Match(para ?? "");
                Ok(run, "tc-id-cell",
                    token.Success && (para ?? "").IndexOf("格外正文", StringComparison.Ordinal) >= 0,
                    para);
                string tokenText = token.Success ? token.Value : "";
                Ok(run, "tc-fold-table-mark",
                    token.Success && tokenText.IndexOf("[in_table]", StringComparison.Ordinal) < 0
                    && tokenText.IndexOf("表", StringComparison.Ordinal) < 0,
                    para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                List<string> ids = MarkerIds(full);
                MatchCollection hits = TocFieldStamp.Marker.Matches(full ?? "");
                string inside = hits.Count >= 2
                    ? (full ?? "").Substring(hits[0].Index, hits[1].Index - hits[0].Index)
                    : "";
                Ok(run, "tc-mark-cell",
                    ids.Count == 2 && ids[0] == ids[1] && ids[0] == (token.Success ? token.Groups[1].Value : "")
                    && inside.IndexOf("第一章", StringComparison.Ordinal) >= 0
                    && inside.IndexOf("格外正文", StringComparison.Ordinal) < 0
                    && (full ?? "").IndexOf("格外正文", StringComparison.Ordinal) >= 0,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunNest(TestRun run, WdSession s)
        {
            if (!Want("tc-id-nest", "tc-mark-nest"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r外层正文。", out path, out id) || !StyleFirst(s, path) || !AddTocInNestedCell(s, path) || !UpdateFields(s, path))
            {
                FailMany(run, "铺嵌套目录失败", "tc-id-nest", "tc-mark-nest");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                int tables = TableCount(s, path);
                Ok(run, "tc-id-nest", Token.IsMatch(para ?? ""), "tables=" + tables + " para=" + para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                List<string> ids = MarkerIds(full);
                MatchCollection hits = TocFieldStamp.Marker.Matches(full ?? "");
                string inside = hits.Count >= 2
                    ? (full ?? "").Substring(hits[0].Index, hits[1].Index - hits[0].Index)
                    : "";
                Ok(run, "tc-mark-nest",
                    ids.Count == 2 && ids[0] == ids[1]
                    && inside.IndexOf("外层正文", StringComparison.Ordinal) < 0
                    && (full ?? "").IndexOf("外层正文", StringComparison.Ordinal) >= 0,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunOrder(TestRun run, WdSession s)
        {
            if (!Want("tc-id-order"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r中间。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !AddTocInCell(s, path) || !UpdateFields(s, path))
            {
                Fail(run, "tc-id-order", "铺顺序稿失败");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                MatchCollection marks = Token.Matches(para ?? "");
                string first = BookmarkAtField(s, path);
                Ok(run, "tc-id-order",
                    marks.Count >= 2 && marks[0].Groups[1].Value == first,
                    "first=" + first + " para=" + para);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunPeriod(TestRun run, WdSession s)
        {
            if (!Want("tc-sent-period"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章。第二节\r你好啊。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !UpdateFields(s, path))
            {
                Fail(run, "tc-sent-period", "铺稿失败");
                return;
            }

            try
            {
                string sent = Sentence(s, id);
                Ok(run, "tc-sent-period",
                    Token.Matches(sent ?? "").Count == 1 && CountOf(sent, "第二节") == 1,
                    sent);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunNotTocField(TestRun run, WdSession s)
        {
            if (!Want("tc-sent-not-toc", "tc-mark-none"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "今天。", out path, out id) || !AddPageField(s, path))
            {
                FailMany(run, "铺页码域失败", "tc-sent-not-toc", "tc-mark-none");
                return;
            }

            try
            {
                string sent = Sentence(s, id);
                Ok(run, "tc-sent-not-toc",
                    (sent ?? "").IndexOf("[目录域", StringComparison.Ordinal) < 0
                    && (sent ?? "").IndexOf("今天", StringComparison.Ordinal) >= 0,
                    sent);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                Ok(run, "tc-mark-none",
                    MarkerIds(full).Count == 0 && (full ?? "").IndexOf("今天", StringComparison.Ordinal) >= 0,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunBothSides(TestRun run, WdSession s)
        {
            if (!Want("tc-sent-both-sides", "tc-mark-both"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "前文。\r第一章标题\r后文。", out path, out id) || !StyleParagraph(s, path, "第一章标题") || !AddTocBefore(s, path, "第一章标题") || !UpdateFields(s, path))
            {
                FailMany(run, "铺稿失败", "tc-sent-both-sides", "tc-mark-both");
                return;
            }

            try
            {
                string sent = Sentence(s, id);
                int token = (sent ?? "").IndexOf("[目录域", StringComparison.Ordinal);
                int before = (sent ?? "").IndexOf("前文", StringComparison.Ordinal);
                int after = (sent ?? "").IndexOf("后文", StringComparison.Ordinal);
                Ok(run, "tc-sent-both-sides",
                    Token.Matches(sent ?? "").Count == 1 && before >= 0 && token > before && after > token,
                    sent);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                MatchCollection hits = TocFieldStamp.Marker.Matches(full ?? "");
                int markAt = hits.Count > 0 ? hits[0].Index : -1;
                int fore = (full ?? "").IndexOf("前文", StringComparison.Ordinal);
                int aft = (full ?? "").IndexOf("后文", StringComparison.Ordinal);
                Ok(run, "tc-mark-both",
                    hits.Count == 2 && fore >= 0 && markAt > fore && aft > hits[1].Index,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCrNext(TestRun run, WdSession s)
        {
            if (!Want("tc-cr-next"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r你好啊。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !UpdateFields(s, path) || !StripResultCr(s, path, true))
            {
                Fail(run, "tc-cr-next", "铺稿失败");
                return;
            }

            try
            {
                int before = ParagraphCount(s, path);
                Paragraph(s, id);
                int after = ParagraphCount(s, path);
                Ok(run, "tc-cr-next", before > 0 && before == after, before + "/" + after);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCrMissing(TestRun run, WdSession s)
        {
            if (!Want("tc-cr-missing", "tc-cr-once", "tc-cr-bookmarked", "tc-fold-mixed"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r你好啊。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !UpdateFields(s, path))
            {
                FailMany(run, "铺缺换行稿失败", "tc-cr-missing", "tc-cr-once", "tc-cr-bookmarked", "tc-fold-mixed");
                return;
            }

            try
            {
                Paragraph(s, id);
                string marked = BookmarkAtField(s, path);
                if (!StripResultCr(s, path, false))
                {
                    FailMany(run, "去不掉结果末尾换行", "tc-cr-missing", "tc-cr-once", "tc-cr-bookmarked", "tc-fold-mixed");
                    return;
                }

                int before = ParagraphCount(s, path);
                string para = Paragraph(s, id);
                int after = ParagraphCount(s, path);
                bool ends = ResultEndsWithCr(s, path);
                Ok(run, "tc-cr-missing",
                    ends && (para ?? "").IndexOf("你好啊", StringComparison.Ordinal) >= 0
                    && Outside(para).IndexOf("第一章标题", StringComparison.Ordinal) >= 0,
                    "ends=" + ends + " " + before + "/" + after + " para=" + para);
                int third = ParagraphCount(s, path);
                Ok(run, "tc-cr-once", after > 0 && after == third, after + "/" + third);
                Ok(run, "tc-cr-bookmarked", BookmarkAtField(s, path) == marked && marked.StartsWith("ewC_", StringComparison.Ordinal), marked);
                Ok(run, "tc-fold-mixed",
                    Token.IsMatch(para ?? "") && Outside(para).IndexOf("第一章标题", StringComparison.Ordinal) >= 0,
                    para);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCrBack(TestRun run, WdSession s)
        {
            if (!Want("tc-cr-back"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "第一章标题\r中间正文。", out path, out id) || !StyleFirst(s, path) || !AddToc(s, path, true) || !AddTocAtEnd(s, path) || !UpdateFields(s, path))
            {
                Fail(run, "tc-cr-back", "铺稿失败");
                return;
            }

            try
            {
                Paragraph(s, id);
                List<string> marks = FieldBookmarks(s, path);
                if (!StripResultCr(s, path, false))
                {
                    Fail(run, "tc-cr-back", "去不掉换行");
                    return;
                }

                Paragraph(s, id);
                List<string> after = FieldBookmarks(s, path);
                Ok(run, "tc-cr-back",
                    marks.Count >= 2 && after.Count >= 2 && marks[0] == after[0] && marks[1] == after[1],
                    string.Join(",", marks) + " -> " + string.Join(",", after));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunPlain(TestRun run, WdSession s)
        {
            if (!Want("tc-plain-none", "tc-plain-style", "tc-mark-style"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "你好啊。", out path, out id))
            {
                FailMany(run, "铺普通稿失败", "tc-plain-none", "tc-plain-style");
                return;
            }

            try
            {
                string sent = Sentence(s, id);
                string para = Paragraph(s, id);
                Ok(run, "tc-plain-none",
                    (sent ?? "").IndexOf("[目录域", StringComparison.Ordinal) < 0
                    && (para ?? "").IndexOf("[目录域", StringComparison.Ordinal) < 0,
                    sent);
            }
            finally
            {
                s.CloseDocByPath(path);
            }

            if (!Open(s, "普通目录行", out path, out id) || !StyleNamed(s, path, "普通目录行", new[] { "目录 1", "TOC 1" }))
            {
                FailMany(run, "铺目录样式失败", "tc-plain-style", "tc-mark-style");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                Ok(run, "tc-plain-style",
                    (para ?? "").IndexOf("普通目录行", StringComparison.Ordinal) >= 0
                    && (para ?? "").IndexOf("[目录域", StringComparison.Ordinal) < 0,
                    para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                Ok(run, "tc-mark-style",
                    MarkerIds(full).Count == 0 && (full ?? "").IndexOf("普通目录行", StringComparison.Ordinal) >= 0,
                    Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunControl(TestRun run, WdSession s)
        {
            if (!Want(
                "tc-ctl-token", "tc-ctl-hide", "tc-ctl-one", "tc-ctl-body", "tc-ctl-tag",
                "tc-ctl-reuse", "tc-ctl-nocr", "tc-ctl-wid", "tc-ctl-inner",
                "tc-mark-ctl", "tc-mark-one"))
            {
                return;
            }

            string path, id;
            if (!Open(s, "占位。", out path, out id) || !InsertXml(s, path, GalleryXml))
            {
                FailMany(run, "铺目录控件失败",
                    "tc-ctl-token", "tc-ctl-hide", "tc-ctl-one", "tc-ctl-body", "tc-ctl-tag",
                    "tc-ctl-reuse", "tc-ctl-nocr", "tc-ctl-wid", "tc-ctl-inner",
                    "tc-mark-ctl", "tc-mark-one");
                return;
            }

            try
            {
                int before = ParagraphCount(s, path);
                string para = Paragraph(s, id);
                Match token = Token.Match(para ?? "");
                string name = token.Success ? token.Groups[1].Value : "";
                string bare = Outside(para);
                Ok(run, "tc-ctl-token", token.Success, para);
                Ok(run, "tc-ctl-hide",
                    bare.IndexOf("键入章标题", StringComparison.Ordinal) < 0
                    && bare.IndexOf("目录", StringComparison.Ordinal) < 0,
                    para);
                Ok(run, "tc-ctl-one", Token.Matches(para ?? "").Count == 1, para);
                Ok(run, "tc-ctl-body", (para ?? "").IndexOf("后文你好啊", StringComparison.Ordinal) >= 0, para);
                Ok(run, "tc-ctl-tag",
                    name.StartsWith("ewC_", StringComparison.Ordinal) && CountBookmarks(s, path, name) == 0
                    && XmlHas(s, path, "w:val=\"" + name + "\""),
                    name);
                string again = Paragraph(s, id);
                Match token2 = Token.Match(again ?? "");
                Ok(run, "tc-ctl-reuse", token2.Success && token2.Groups[1].Value == name, again);
                Ok(run, "tc-ctl-nocr", before > 0 && before == ParagraphCount(s, path), before.ToString());
                Ok(run, "tc-ctl-wid", name != "1051109989", name);
                Ok(run, "tc-ctl-inner", Token.Matches(para ?? "").Count == 1, para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                List<string> ids = MarkerIds(full);
                MatchCollection hits = TocFieldStamp.Marker.Matches(full ?? "");
                string inside = hits.Count >= 2
                    ? (full ?? "").Substring(hits[0].Index, hits[1].Index - hits[0].Index)
                    : "";
                Ok(run, "tc-mark-ctl",
                    ids.Count == 2 && ids[0] == ids[1] && ids[0] == name
                    && inside.IndexOf("键入章标题", StringComparison.Ordinal) >= 0
                    && inside.IndexOf("后文你好啊", StringComparison.Ordinal) < 0,
                    string.Join(",", ids));
                Ok(run, "tc-mark-one", ids.Count == 2, "n=" + ids.Count + " " + Brief(full));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunControlField(TestRun run, WdSession s)
        {
            if (!Want("tc-ctl-field", "tc-mark-field"))
            {
                return;
            }

            string path, id;
            const string xml = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/_rels/.rels\" pkg:contentType=\"application/vnd.openxmlformats-package.relationships+xml\"><pkg:xmlData><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:sdt><w:sdtPr><w:id w:val=\"99\"/><w:docPartObj><w:docPartGallery w:val=\"Table of Contents\"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> TOC \\o \"1-3\" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>目录条目</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p></w:sdtContent></w:sdt><w:p><w:r><w:t>后文。</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
            if (!Open(s, "占位。", out path, out id) || !InsertXml(s, path, xml))
            {
                FailMany(run, "铺控件内目录域失败", "tc-ctl-field", "tc-mark-field");
                return;
            }

            try
            {
                string para = Paragraph(s, id);
                MatchCollection marks = Token.Matches(para ?? "");
                string name = marks.Count > 0 ? marks[0].Groups[1].Value : "";
                Ok(run, "tc-ctl-field",
                    marks.Count == 1 && CountBookmarks(s, path, name) == 1,
                    "n=" + marks.Count + " marks=" + CountBookmarks(s, path, name) + " para=" + para);
                string full = "";
                string agent = "";
                ReadDisplays(s, path, ref full, ref agent);
                List<string> ids = MarkerIds(full);
                Ok(run, "tc-mark-field",
                    ids.Count == 2 && ids[0] == ids[1] && ids[0] == name,
                    string.Join(",", ids) + " name=" + name);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
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

        private static string Sentence(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "sentence");
            return got != null && got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string Paragraph(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id, "paragraph");
            return got != null && got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static string Chunk(WdSession s, string id)
        {
            ToolResult got = DesktopToolPipe.Call(
                "F_get_curr_doc_chunk_paragraph_display_content",
                new Dictionary<string, object> { ["channel_id"] = id, ["chunk_idx"] = 1 });
            if (got == null)
            {
                return "没有回包";
            }

            if (!got.Success)
            {
                return got.Error ?? "失败";
            }

            return WdSession.Str(got, "paragraph_display_content");
        }

        private static bool StyleFirst(WdSession s, string path)
        {
            return StyleParagraph(s, path, null);
        }

        private static bool StyleParagraph(WdSession s, string path, string contains)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => StyleLate(doc, contains, HeadingNames))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, contains);
                    object style = FirstStyle(doc, HeadingNames);
                    if (para == null || style == null)
                    {
                        return false;
                    }

                    para.set_Style(ref style);
                    return true;
                });
        }

        private static bool StyleNamed(WdSession s, string path, string contains, string[] names)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => StyleLate(doc, contains, names))
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, contains);
                    object style = FirstStyle(doc, names);
                    if (para == null || style == null)
                    {
                        return false;
                    }

                    para.set_Style(ref style);
                    return true;
                });
        }

        private static bool StyleLate(object doc, string contains, string[] names)
        {
            object para = FindParaLate(doc, contains);
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

            if (para == null || style == null)
            {
                return false;
            }

            WdSession.ComSet(para, "Style", style);
            return true;
        }

        private static object FirstStyle(Word.Document doc, string[] names)
        {
            foreach (string name in names)
            {
                try
                {
                    return doc.Styles[name];
                }
                catch (Exception)
                {
                }
            }

            return null;
        }

        private static Word.Paragraph FindPara(Word.Document doc, string contains)
        {
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                string text = para.Range.Text ?? "";
                if (contains == null || text.IndexOf(contains, StringComparison.Ordinal) >= 0)
                {
                    return para;
                }
            }

            return null;
        }

        private static object FindParaLate(object doc, string contains)
        {
            object paras = WdSession.ComGet(doc, "Paragraphs");
            foreach (object para in (System.Collections.IEnumerable)paras)
            {
                object range = WdSession.ComGet(para, "Range");
                string text = Convert.ToString(WdSession.ComGet(range, "Text")) ?? "";
                if (contains == null || text.IndexOf(contains, StringComparison.Ordinal) >= 0)
                {
                    return para;
                }
            }

            return null;
        }

        private static bool AddToc(WdSession s, string path, bool atStart)
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
                    if (atStart)
                    {
                        doc.Range(0, 0).InsertParagraphBefore();
                    }

                    Word.Range range = doc.Paragraphs[1].Range.Duplicate;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    doc.Fields.Add(range, Word.WdFieldType.wdFieldTOC, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                });
        }

        private static bool AddTocBefore(WdSession s, string path, string contains)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object para = FindParaLate(doc, contains);
                    if (para == null)
                    {
                        return false;
                    }

                    object range = WdSession.ComGet(para, "Range");
                    WdSession.ComCall(range, "Collapse", 1);
                    object fields = WdSession.ComGet(doc, "Fields");
                    WdSession.ComCall(fields, "Add", range, 13, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = FindPara(doc, contains);
                    if (para == null)
                    {
                        return false;
                    }

                    Word.Range range = para.Range.Duplicate;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    doc.Fields.Add(range, Word.WdFieldType.wdFieldTOC, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                });
        }

        private static bool AddTocAtEnd(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object content = WdSession.ComGet(doc, "Content");
                    int end = Convert.ToInt32(WdSession.ComGet(content, "End"));
                    object range = WdSession.ComCall(doc, "Range", end - 1, end - 1);
                    object fields = WdSession.ComGet(doc, "Fields");
                    WdSession.ComCall(fields, "Add", range, 13, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    int end = doc.Content.End;
                    Word.Range range = doc.Range(end - 1, end - 1);
                    doc.Fields.Add(range, Word.WdFieldType.wdFieldTOC, "\\o \"1-3\" \\h \\z \\u", false);
                    return true;
                });
        }

        private static bool AddTocInCell(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => AddCellLate(doc, false))
                : s.OnWordDoc(path, doc => AddCell(doc, false));
        }

        private static bool AddTocInNestedCell(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => AddCellLate(doc, true))
                : s.OnWordDoc(path, doc => AddCell(doc, true));
        }

        private static bool AddCell(Word.Document doc, bool nest)
        {
            Word.Table outer = doc.Tables.Add(doc.Range(doc.Content.End - 1, doc.Content.End - 1), 1, 1);
            Word.Range cell = outer.Cell(1, 1).Range;
            if (nest)
            {
                Word.Table inner = doc.Tables.Add(cell, 1, 1);
                cell = inner.Cell(1, 1).Range;
            }

            cell.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            doc.Fields.Add(cell, Word.WdFieldType.wdFieldTOC, "\\o \"1-3\" \\h \\z \\u", false);
            return true;
        }

        private static bool AddCellLate(object doc, bool nest)
        {
            object content = WdSession.ComGet(doc, "Content");
            int end = Convert.ToInt32(WdSession.ComGet(content, "End"));
            object range = WdSession.ComCall(doc, "Range", end - 1, end - 1);
            object tables = WdSession.ComGet(doc, "Tables");
            object outer = WdSession.ComCall(tables, "Add", range, 1, 1);
            object cell = WdSession.ComCall(outer, "Cell", 1, 1);
            object cellRange = WdSession.ComGet(cell, "Range");
            if (nest)
            {
                object inner = WdSession.ComCall(tables, "Add", cellRange, 1, 1);
                cell = WdSession.ComCall(inner, "Cell", 1, 1);
                cellRange = WdSession.ComGet(cell, "Range");
            }

            WdSession.ComCall(cellRange, "Collapse", 1);
            object fields = WdSession.ComGet(doc, "Fields");
            WdSession.ComCall(fields, "Add", cellRange, 13, "\\o \"1-3\" \\h \\z \\u", false);
            return true;
        }

        private static bool AddPageField(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object range = WdSession.ComGet(doc, "Range");
                    WdSession.ComCall(range, "Collapse", 1);
                    object fields = WdSession.ComGet(doc, "Fields");
                    WdSession.ComCall(fields, "Add", range, 33, "", false);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Range range = doc.Range(0, 0);
                    doc.Fields.Add(range, Word.WdFieldType.wdFieldPage, "", false);
                    return true;
                });
        }

        private static bool UpdateFields(WdSession s, string path)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    WdSession.ComCall(WdSession.ComGet(doc, "Fields"), "Update");
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    doc.Fields.Update();
                    return true;
                });
        }

        private static string PlantNote;

        private static bool PlantTocAtStart(WdSession s, string path)
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

        private static bool PlantTocOn(WdSession s, string path, string contains)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object paras = WdSession.ComGet(doc, "Paragraphs");
                    object para = null;
                    foreach (object item in (System.Collections.IEnumerable)paras)
                    {
                        string text = Convert.ToString(WdSession.ComGet(WdSession.ComGet(item, "Range"), "Text")) ?? "";
                        if (text.IndexOf(contains, StringComparison.Ordinal) >= 0 && text.IndexOf('\t') < 0)
                        {
                            para = item;
                        }
                    }

                    if (para == null)
                    {
                        return false;
                    }

                    object range = WdSession.ComGet(para, "Range");
                    WdSession.ComCall(range, "Collapse", 1);
                    object bookmarks = WdSession.ComGet(doc, "Bookmarks");
                    WdSession.ComCall(bookmarks, "Add", "_Toc245561084", range);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    Word.Paragraph para = null;
                    foreach (Word.Paragraph item in doc.Paragraphs)
                    {
                        string text = item.Range.Text ?? "";
                        if (text.IndexOf(contains, StringComparison.Ordinal) >= 0
                            && text.IndexOf('\t') < 0)
                        {
                            para = item;
                        }
                    }

                    if (para == null)
                    {
                        return false;
                    }

                    Word.Range range = para.Range.Duplicate;
                    range.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                    object anchor = range;
                    doc.Bookmarks.Add("_Toc245561084", ref anchor);
                    return true;
                });
        }

        private static bool InsertXml(WdSession s, string path, string xml)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    WdSession.ComCall(WdSession.ComGet(doc, "Content"), "InsertXML", xml);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    doc.Content.InsertXML(xml);
                    return true;
                });
        }

        private static string BookmarkAtField(WdSession s, string path)
        {
            List<string> names = FieldBookmarks(s, path);
            return names.Count == 0 ? "" : names[0];
        }

        private static List<string> FieldBookmarks(WdSession s, string path)
        {
            var names = new List<string>();
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    names.AddRange(FieldBookmarksLate(doc));
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    foreach (Word.Field field in doc.Fields)
                    {
                        if ((int)field.Type != (int)Word.WdFieldType.wdFieldTOC)
                        {
                            continue;
                        }

                        Word.Range point = field.Code.Duplicate;
                        point.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                        foreach (Word.Bookmark bookmark in doc.Bookmarks)
                        {
                            if (bookmark.Range != null && bookmark.Range.Start == point.Start
                                && (bookmark.Name ?? "").StartsWith("ewC_", StringComparison.Ordinal))
                            {
                                names.Add(bookmark.Name);
                            }
                        }
                    }

                    return true;
                });
            }

            return names;
        }

        private static List<string> FieldBookmarksLate(object doc)
        {
            var names = new List<string>();
            object fields = WdSession.ComGet(doc, "Fields");
            object bookmarks = WdSession.ComGet(doc, "Bookmarks");
            foreach (object field in (System.Collections.IEnumerable)fields)
            {
                object type = WdSession.ComGet(field, "Type");
                if (type == null || Convert.ToInt32(type) != 13)
                {
                    continue;
                }

                object code = WdSession.ComGet(field, "Code");
                WdSession.ComCall(code, "Collapse", 1);
                int start = Convert.ToInt32(WdSession.ComGet(code, "Start"));
                foreach (object bookmark in (System.Collections.IEnumerable)bookmarks)
                {
                    object range = WdSession.ComGet(bookmark, "Range");
                    int at = Convert.ToInt32(WdSession.ComGet(range, "Start"));
                    string name = Convert.ToString(WdSession.ComGet(bookmark, "Name")) ?? "";
                    if (at == start && name.StartsWith("ewC_", StringComparison.Ordinal))
                    {
                        names.Add(name);
                    }
                }
            }

            return names;
        }

        private static int CountBookmarks(WdSession s, string path, string name)
        {
            int count = 0;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object bookmarks = WdSession.ComGet(doc, "Bookmarks");
                    foreach (object bookmark in (System.Collections.IEnumerable)bookmarks)
                    {
                        if (string.Equals(Convert.ToString(WdSession.ComGet(bookmark, "Name")), name, StringComparison.Ordinal))
                        {
                            count++;
                        }
                    }

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

        private static bool BookmarkExists(WdSession s, string path, string name)
        {
            bool found = false;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object exists = WdSession.ComCall(WdSession.ComGet(doc, "Bookmarks"), "Exists", name);
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

        private static string BookmarkNames(WdSession s, string path)
        {
            var names = new List<string>();
            s.OnWordDoc(path, doc =>
            {
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    names.Add(bookmark.Name);
                }

                return true;
            });
            return string.Join(",", names);
        }

        private static int ParagraphCount(WdSession s, string path)
        {
            int count = -1;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    count = Convert.ToInt32(WdSession.ComGet(WdSession.ComGet(doc, "Paragraphs"), "Count"));
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    count = doc.Paragraphs.Count;
                    return true;
                });
            return ok ? count : -1;
        }

        private static int TableCount(WdSession s, string path)
        {
            int count = -1;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    count = Convert.ToInt32(WdSession.ComGet(WdSession.ComGet(doc, "Tables"), "Count"));
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    count = doc.Tables.Count;
                    return true;
                });
            return ok ? count : -1;
        }

        private static bool BodyParagraphAfterToc(WdSession s, string path, string text)
        {
            int toc = -1;
            int body = -1;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    object paras = WdSession.ComGet(doc, "Paragraphs");
                    int n = Convert.ToInt32(WdSession.ComGet(paras, "Count"));
                    for (int i = 1; i <= n; i++)
                    {
                        object para = WdSession.ComCall(paras, "Item", i);
                        string value = Convert.ToString(WdSession.ComGet(WdSession.ComGet(para, "Range"), "Text")) ?? "";
                        if (toc < 0 && value.IndexOf("\t", StringComparison.Ordinal) >= 0)
                        {
                            toc = i;
                        }

                        if (value.IndexOf(text, StringComparison.Ordinal) >= 0)
                        {
                            body = i;
                        }
                    }

                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    int i = 0;
                    foreach (Word.Paragraph para in doc.Paragraphs)
                    {
                        i++;
                        string value = para.Range.Text ?? "";
                        if (toc < 0 && value.IndexOf("\t", StringComparison.Ordinal) >= 0)
                        {
                            toc = i;
                        }

                        if (value.IndexOf(text, StringComparison.Ordinal) >= 0)
                        {
                            body = i;
                        }
                    }

                    return true;
                });
            return ok && toc > 0 && body > toc;
        }

        private static bool StripResultCr(WdSession s, string path, bool keepFollowingCr)
        {
            return s.IsWps
                ? s.TryWpsDoc(path, doc => StripLate(doc, keepFollowingCr))
                : s.OnWordDoc(path, doc =>
                {
                    var fields = new List<Word.Field>();
                    foreach (Word.Field field in doc.Fields)
                    {
                        if ((int)field.Type == (int)Word.WdFieldType.wdFieldTOC)
                        {
                            fields.Add(field);
                        }
                    }

                    for (int i = fields.Count - 1; i >= 0; i--)
                    {
                        Word.Range result = fields[i].Result;
                        string text = result.Text ?? "";
                        if (text.Length == 0 || text[text.Length - 1] != '\r')
                        {
                            continue;
                        }

                        if (keepFollowingCr)
                        {
                            Word.Range next = doc.Range(result.End, Math.Min(doc.Content.End, result.End + 1));
                            if ((next.Text ?? "").IndexOf('\r') != 0)
                            {
                                doc.Range(result.End, result.End).InsertAfter("\r");
                            }
                        }

                        doc.Range(result.End - 1, result.End).Delete();
                    }

                    return true;
                });
        }

        private static bool StripLate(object doc, bool keepFollowingCr)
        {
            object fieldsObj = WdSession.ComGet(doc, "Fields");
            var fields = new List<object>();
            foreach (object field in (System.Collections.IEnumerable)fieldsObj)
            {
                object type = WdSession.ComGet(field, "Type");
                if (type != null && Convert.ToInt32(type) == 13)
                {
                    fields.Add(field);
                }
            }

            object content = WdSession.ComGet(doc, "Content");
            int contentEnd = Convert.ToInt32(WdSession.ComGet(content, "End"));
            for (int i = fields.Count - 1; i >= 0; i--)
            {
                object result = WdSession.ComGet(fields[i], "Result");
                string text = Convert.ToString(WdSession.ComGet(result, "Text")) ?? "";
                if (text.Length == 0 || text[text.Length - 1] != '\r')
                {
                    continue;
                }

                int end = Convert.ToInt32(WdSession.ComGet(result, "End"));
                if (keepFollowingCr)
                {
                    object next = WdSession.ComCall(doc, "Range", end, Math.Min(contentEnd, end + 1));
                    string nextText = Convert.ToString(WdSession.ComGet(next, "Text")) ?? "";
                    if (nextText.IndexOf('\r') != 0)
                    {
                        object at = WdSession.ComCall(doc, "Range", end, end);
                        WdSession.ComCall(at, "InsertAfter", "\r");
                    }
                }

                object mark = WdSession.ComCall(doc, "Range", end - 1, end);
                WdSession.ComCall(mark, "Delete");
            }

            return true;
        }

        private static bool ResultEndsWithCr(WdSession s, string path)
        {
            bool ends = false;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    foreach (object field in (System.Collections.IEnumerable)WdSession.ComGet(doc, "Fields"))
                    {
                        object type = WdSession.ComGet(field, "Type");
                        if (type == null || Convert.ToInt32(type) != 13)
                        {
                            continue;
                        }

                        string text = Convert.ToString(WdSession.ComGet(WdSession.ComGet(field, "Result"), "Text")) ?? "";
                        ends = text.Length > 0 && text[text.Length - 1] == '\r';
                        return true;
                    }

                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    foreach (Word.Field field in doc.Fields)
                    {
                        if ((int)field.Type != (int)Word.WdFieldType.wdFieldTOC)
                        {
                            continue;
                        }

                        string text = field.Result.Text ?? "";
                        ends = text.Length > 0 && text[text.Length - 1] == '\r';
                        return true;
                    }

                    return true;
                });
            return ok && ends;
        }

        private static void ReadDisplays(WdSession s, string path, ref string full, ref string agent)
        {
            string fullLocal = "";
            string agentLocal = "";
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    Word.Document wordDoc = doc as Word.Document;
                    if (wordDoc == null)
                    {
                        return false;
                    }

                    WordDocumentExtractor.ProcessDocument(wordDoc, ProcessDocumentOptions.ForProcessActions("locate"));
                    fullLocal = string.Join("", DocumentState.ChunkParagraphDisplayContents);
                    agentLocal = string.Join("", DocumentState.ChunkParagraphAgentDisplays);
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    WordDocumentExtractor.ProcessDocument(doc, ProcessDocumentOptions.ForProcessActions("locate"));
                    fullLocal = string.Join("", DocumentState.ChunkParagraphDisplayContents);
                    agentLocal = string.Join("", DocumentState.ChunkParagraphAgentDisplays);
                    return true;
                });
            }

            full = fullLocal;
            agent = agentLocal;
        }

        private static string FirstInteriorParagraph(WdSession s, string path)
        {
            string code = "";
            ReadInterior(s, path, ref code, false);
            return code;
        }

        private static string FirstInteriorSentence(WdSession s, string path)
        {
            string code = "";
            ReadInterior(s, path, ref code, true);
            return code;
        }

        private static void ReadInterior(WdSession s, string path, ref string code, bool sentence)
        {
            string found = "";
            if (s.IsWps)
            {
                s.TryWpsDoc(path, doc =>
                {
                    Word.Document wordDoc = doc as Word.Document;
                    if (wordDoc == null)
                    {
                        return false;
                    }

                    WordDocumentExtractor.ProcessDocument(wordDoc, ProcessDocumentOptions.ForProcessActions("locate"));
                    found = sentence
                        ? TocFieldStamp.FirstInteriorSentence() ?? ""
                        : TocFieldStamp.FirstInteriorParagraph() ?? "";
                    return true;
                });
            }
            else
            {
                s.OnWordDoc(path, doc =>
                {
                    WordDocumentExtractor.ProcessDocument(doc, ProcessDocumentOptions.ForProcessActions("locate"));
                    found = sentence
                        ? TocFieldStamp.FirstInteriorSentence() ?? ""
                        : TocFieldStamp.FirstInteriorParagraph() ?? "";
                    return true;
                });
            }

            code = found;
        }

        private static string Replace(WdSession s, string id, string code, string neu)
        {
            if (string.IsNullOrEmpty(code))
            {
                return "没有码";
            }

            ToolResult result = s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = code },
                        ["new"] = neu,
                        ["format"] = new Dictionary<string, object>
                        {
                            ["char_format"] = new Dictionary<string, object>
                            {
                                ["font_name"] = "宋体",
                                ["font_size"] = 12,
                                ["is_bold"] = false
                            }
                        }
                    }
                }
            });
            if (result == null)
            {
                return "没有回包";
            }

            return result.Success ? "" : (result.Error ?? "");
        }

        private static string Align(WdSession s, string id, string code)
        {
            if (string.IsNullOrEmpty(code))
            {
                return "没有段码";
            }

            ToolResult bad = s.ApplyParagraphFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["target_paragraph_codes"] = code,
                ["para_format"] = new Dictionary<string, object> { ["alignment"] = "center" }
            });
            return bad == null ? "没有回包" : (bad.Success ? "" : (bad.Error ?? ""));
        }

        private static string StyleInterior(WdSession s, string path)
        {
            string error = "";
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    Word.Document wordDoc = doc as Word.Document;
                    if (wordDoc == null)
                    {
                        error = "打不开稿";
                        return false;
                    }

                    WordDocumentExtractor.ProcessDocument(wordDoc, ProcessDocumentOptions.ForProcessActions("locate"));
                    string code = TocFieldStamp.FirstInteriorSentence() ?? "";
                    error = ApplyStyle(wordDoc, code);
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    WordDocumentExtractor.ProcessDocument(doc, ProcessDocumentOptions.ForProcessActions("locate"));
                    error = ApplyStyle(doc, TocFieldStamp.FirstInteriorSentence() ?? "");
                    return true;
                });
            return ok ? error : "没打开稿";
        }

        private static string ApplyStyle(Word.Document doc, string code)
        {
            if (code.Length == 0)
            {
                return "没有域内句子码";
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
            return result != null && result.ContainsKey("error") ? Convert.ToString(result["error"]) : "";
        }

        private static JArray Spans(WdSession s, string id)
        {
            return Arr(s.DocumentStructure(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["action"] = "list_spans"
            }));
        }

        private static JArray Styles(WdSession s, string id)
        {
            return Arr(s.DocumentStructure(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["action"] = "list_styles"
            }), "styles");
        }

        private static JArray Arr(ToolResult result, string key = "spans")
        {
            if (result == null || result.Data == null)
            {
                return new JArray();
            }

            try
            {
                JObject data = JObject.Parse(Newtonsoft.Json.JsonConvert.SerializeObject(result.Data));
                return data[key] as JArray ?? new JArray();
            }
            catch (Exception)
            {
                return new JArray();
            }
        }

        private static bool SpanHas(JArray spans, string text)
        {
            foreach (JToken span in spans)
            {
                if ((span.ToString() ?? "").IndexOf(text, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static JToken SpanByCode(JArray spans, string code)
        {
            if (!code.StartsWith("S_", StringComparison.Ordinal))
            {
                return null;
            }

            foreach (JToken span in spans)
            {
                if ((span["target_codes"] ?? "").ToString().IndexOf(code, StringComparison.Ordinal) >= 0)
                {
                    return span;
                }
            }

            return null;
        }

        private static JToken FindStyle(JArray styles, string name)
        {
            foreach (JToken style in styles)
            {
                string structure = (style["structure"] ?? "").ToString();
                string label = (style["name"] ?? "").ToString();
                if (string.Equals(structure, name, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(label, name, StringComparison.OrdinalIgnoreCase)
                    || (style.ToString() ?? "").IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return style;
                }
            }

            return null;
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

            int dash = body.IndexOf("-start", mark, StringComparison.Ordinal);
            if (dash < 0)
            {
                return "";
            }

            return body.Substring(mark + 1, dash - mark - 1);
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
            int end = mark < 0 ? -1 : (body ?? "").IndexOf("[S_" + code.Substring(2) + "-end]", mark, StringComparison.Ordinal);
            if (end < 0)
            {
                return false;
            }

            return (body ?? "").Substring(mark, end - mark).IndexOf(extra, StringComparison.Ordinal) >= 0;
        }

        private static bool InsideToken(string body, string text)
        {
            int at = (body ?? "").IndexOf(text, StringComparison.Ordinal);
            if (at < 0)
            {
                return false;
            }

            int token = (body ?? "").LastIndexOf("[目录域", at, StringComparison.Ordinal);
            int close = token < 0 ? -1 : (body ?? "").IndexOf(']', token);
            return close >= 0 && close < at && at < close + 8;
        }

        private static string Outside(string display)
        {
            return Token.Replace(display ?? "", "");
        }

        private static int CountOf(string text, string needle)
        {
            int count = 0;
            int at = 0;
            while (!string.IsNullOrEmpty(text) && (at = text.IndexOf(needle, at, StringComparison.Ordinal)) >= 0)
            {
                count++;
                at += needle.Length;
            }

            return count;
        }

        private static string Brief(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Length > 180 ? text.Substring(0, 180) : text;
        }

        private static string Dump(JToken token)
        {
            string text = token == null ? "" : token.ToString();
            return text.Length > 240 ? text.Substring(0, 240) : text;
        }

        private static bool XmlHas(WdSession s, string path, string needle)
        {
            bool has = false;
            bool ok = s.IsWps
                ? s.TryWpsDoc(path, doc =>
                {
                    string xml = Convert.ToString(WdSession.ComGet(WdSession.ComGet(doc, "Content"), "WordOpenXML")) ?? "";
                    has = xml.IndexOf(needle, StringComparison.Ordinal) >= 0;
                    return true;
                })
                : s.OnWordDoc(path, doc =>
                {
                    string xml = doc.Content.WordOpenXML ?? "";
                    has = xml.IndexOf(needle, StringComparison.Ordinal) >= 0;
                    return true;
                });
            return ok && has;
        }
    }
}
