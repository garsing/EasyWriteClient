using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>编码时原文消歧 COM 短套。对照 docs/测试/word测试/编码时原文消歧。</summary>
    internal static class WuCases
    {
        private const string Unique = "独一份句子。";
        private const string Merge = "甲。乙。甲。丙。";
        private const string Keep = "头句独一份内容。短。尾句独一份。短。";
        private const string Tail = "独一份前句。短。";
        private const string Dup = "同句。";
        private const string Splitter = "隔开格。";
        private const string CellSame = "格子短句。";
        private const string FillMerged = "改后的甲乙。";
        private const string FillGeo = "填进左格。";
        private const string Title = "实验操作方法和步骤";
        private const string TailText = "尾巴。";

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            run.TotalTarget = 19;
            var s = new WdSession(host);
            RunTyped(run, s, "wu-unique", Unique, body =>
                HasS(body) && body.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0 && !HasGeo(body));
            RunTyped(run, s, "wu-merge-fwd", Merge, body =>
                body.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0
                && CountSStarts(body) >= 2);
            RunTyped(run, s, "wu-keep-unique", Keep, body =>
                body.IndexOf("头句独一份内容。", StringComparison.Ordinal) >= 0
                && !SpanContains(body, "头句独一份内容。", "短。"));
            RunTyped(run, s, "wu-tail-join", Tail + "\r" + "短。", body =>
                body.IndexOf("独一份前句。", StringComparison.Ordinal) >= 0
                && body.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0);

            RunDupSplit(run, s);
            RunCells(run, s);
            RunCr9(run, s);
            RunOverlap(run, s);
            RunVacuum(run, s);
            RunLongCell(run, s);
            RunWrites(run, s);
            RunParaCr(run, s);

            string wordDetail, wpsDetail;
            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
            Ok(run, "wu-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        /// <summary>
        /// 造 2×2（有字/真空/有字/真空），看段 display 空格是否有 P_，再用 P_ 套段、用几何码填格。
        /// </summary>
        public static int RunParaVacuumOp(string host)
        {
            var run = new TestRun();
            run.TotalTarget = 7;
            var s = new WdSession(host);
            string path;
            string ch;
            if (!TryOpen(run, s, out path, out ch, "wu-pvac-open"))
            {
                Console.WriteLine("正式用例 通过 " + run.CasesPassed
                    + "  失败 " + run.CasesFailed);
                return 1;
            }

            try
            {
                if (!s.InsertTable(path, 2, 2, new[] { Title, "", "备注", "\r\r" }))
                {
                    FailMany(run, "COM 真空表失败", "wu-pvac-geo", "wu-pvac-empty-p",
                        "wu-pvac-fmt-ctx", "wu-pvac-apply", "wu-pvac-com-align",
                        "wu-pvac-write-p", "wu-pvac-write-geo");
                    return 1;
                }

                string sent = ReadBody(s, ch);
                var geos = new List<string>();
                foreach (Match m in Regex.Matches(sent, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+"))
                {
                    if (!geos.Contains(m.Value))
                    {
                        geos.Add(m.Value);
                    }
                }

                Console.WriteLine("句 display 几何码 " + geos.Count + ": " + string.Join(", ", geos));
                Console.WriteLine("句预览 " + Preview(sent));
                Ok(run, "wu-pvac-geo", geos.Count >= 2, "真空几何码=" + geos.Count);

                ToolResult paraGot = s.GetContent(ch, "paragraph");
                string para = paraGot.Success ? WdSession.Str(paraGot, "document_content") : "";
                Console.WriteLine("段预览 " + Preview(para));
                List<string> emptyPs = CollectEmptyCellParagraphCodes(para);
                Console.WriteLine("空 td 内 P_ " + emptyPs.Count + ": " + string.Join(", ", emptyPs));
                DumpTdBodies(para);
                DumpParagraphBlocks(para);
                Ok(run, "wu-pvac-empty-p", emptyPs.Count >= 1, "空 td 内 P_=" + emptyPs.Count);

                string pCode = emptyPs.Count > 0 ? emptyPs[0] : "";
                ToolResult fmt = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["target_paragraph_codes"] = pCode
                });
                Console.WriteLine("get_format_context(P_) success=" + fmt.Success
                    + " err=" + TrimBlob(Err(fmt)));
                Ok(run, "wu-pvac-fmt-ctx",
                    fmt.Success && !Contains(Err(fmt), "empty_cell_no_format"),
                    fmt.Success ? "P_ 取段格式成功" : TrimBlob(Err(fmt)));

                ToolResult apply = s.ApplyParagraphFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["target_paragraph_codes"] = pCode,
                    ["para_format"] = new Dictionary<string, object> { ["alignment"] = "center" }
                });
                Console.WriteLine("apply_paragraph_format(P_) success=" + apply.Success
                    + " err=" + TrimBlob(Err(apply)));
                Ok(run, "wu-pvac-apply", apply.Success, apply.Success ? "套段成功" : TrimBlob(Err(apply)));

                int align = -1;
                s.OnWordDoc(path, doc =>
                {
                    Word.Range vacuum = doc.Tables[1].Cell(1, 2).Range;
                    Word.Range crCell = doc.Tables[1].Cell(2, 2).Range;
                    Console.WriteLine("COM 真空格(1,2) paras=" + vacuum.Paragraphs.Count
                        + " Text=" + QuoteCr(vacuum.Text)
                        + " align0=" + (int)vacuum.Paragraphs[1].Range.ParagraphFormat.Alignment);
                    Console.WriteLine("COM 空段格(2,2) paras=" + crCell.Paragraphs.Count
                        + " Text=" + QuoteCr(crCell.Text)
                        + " align0=" + (int)crCell.Paragraphs[1].Range.ParagraphFormat.Alignment);
                    align = (int)crCell.Paragraphs[1].Range.ParagraphFormat.Alignment;
                    return true;
                });
                Ok(run, "wu-pvac-com-align",
                    apply.Success && align == (int)Word.WdParagraphAlignment.wdAlignParagraphCenter,
                    "align=" + align);

                string geo = geos.Count > 0 ? geos[0] : "";
                ToolResult writeG = Replace(s, ch, geo, FillGeo);
                string after = ReadBody(s, ch);
                Console.WriteLine("replace 几何码 success=" + writeG.Success
                    + " 稿含填空=" + (after.IndexOf(FillGeo, StringComparison.Ordinal) >= 0)
                    + " err=" + TrimBlob(Err(writeG)));
                Ok(run, "wu-pvac-write-geo",
                    writeG.Success && after.IndexOf(FillGeo, StringComparison.Ordinal) >= 0,
                    writeG.Success ? "几何码写入真空格" : TrimBlob(Err(writeG)));

                ToolResult writeP = Replace(s, ch, pCode, "不该用句码字段写P。");
                Console.WriteLine("replace codes=P_ success=" + writeP.Success
                    + " err=" + TrimBlob(Err(writeP)));
                Ok(run, "wu-pvac-write-p",
                    !writeP.Success,
                    writeP.Success ? "不该用 codes 吃 P_" : TrimBlob(Err(writeP)));
            }
            finally
            {
                s.CloseDocByPath(path);
            }

            Console.WriteLine("正式用例 通过 " + run.CasesPassed
                + "  失败 " + run.CasesFailed
                + "  跳过 " + run.CasesSkipped);
            return run.CasesFailed == 0 ? 0 : 1;
        }

        private static void RunTyped(TestRun run, WdSession s, string id, string text, Func<string, bool> ok)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, id))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, text))
                {
                    Ok(run, id, false, "COM 写字失败");
                    return;
                }

                string body = ReadBody(s, ch);
                Ok(run, id, ok(body), Preview(body));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunDupSplit(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-no-cross-seg", "wu-d22-dup", "wu-write-d22", "wu-fmt-d22"))
            {
                return;
            }

            try
            {
                if (!WriteDupAroundTable(s, path))
                {
                    FailMany(run, "COM 写两段同句失败", "wu-no-cross-seg", "wu-d22-dup", "wu-write-d22", "wu-fmt-d22");
                    return;
                }

                string body = ReadBody(s, ch);
                Ok(run, "wu-no-cross-seg",
                    body.IndexOf("乙。", StringComparison.Ordinal) < 0
                    || body.IndexOf("歧义无法操作", StringComparison.Ordinal) >= 0,
                    Preview(body));
                bool d22 = body.IndexOf("-start 歧义无法操作]", StringComparison.Ordinal) >= 0;
                Ok(run, "wu-d22-dup", d22, Preview(body));
                string d22Code = D22Containing(body, Dup);

                ToolResult write = Replace(s, ch, d22Code, "不该写入。");
                string writeBlob = Err(write);
                Ok(run, "wu-write-d22",
                    !write.Success
                    && (Contains(writeBlob, "ambiguous_unusable_sentence")
                        || Contains(writeBlob, "歧义无法操作"))
                    && ReadBody(s, ch).IndexOf("不该写入。", StringComparison.Ordinal) < 0,
                    write.Success ? "D22 不该写入" : writeBlob);

                ToolResult fmt = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["target_codes"] = d22Code
                });
                Ok(run, "wu-fmt-d22",
                    !fmt.Success && (Contains(fmt.Error, "ambiguous_unusable_sentence")
                        || Contains(fmt.Error, "歧义无法操作")),
                    fmt.Success ? "D22 取格式不该成功" : (fmt.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCells(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-cell-same", "wu-write-geo"))
            {
                return;
            }

            try
            {
                ToolResult created = s.CreateTableFromXml(CreateArgs(ch, WriteXml(TwoCellXml(CellSame, CellSame))));
                if (!created.Success)
                {
                    FailMany(run, created.Error ?? "建两格失败", "wu-cell-same", "wu-write-geo");
                    return;
                }

                string body = ReadBody(s, ch);
                bool geo = CountGeoStarts(body) >= 2;
                Ok(run, "wu-cell-same", geo, Preview(body));
                string code = FirstGeo(body);
                ToolResult rep = Replace(s, ch, code, FillGeo);
                string after = ReadBody(s, ch);
                Ok(run, "wu-write-geo",
                    rep.Success && after.IndexOf(FillGeo, StringComparison.Ordinal) >= 0,
                    rep.Success ? Preview(after) : Err(rep));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCr9(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-cell-cr9"))
            {
                return;
            }

            try
            {
                string left = TailText + new string('\r', 8);
                string right = new string('\r', 9);
                if (!s.InsertTable(path, 1, 2, new[] { left, right }))
                {
                    Ok(run, "wu-cell-cr9", false, "COM 建粘段表失败");
                    return;
                }

                string body = ReadBody(s, ch);
                bool leftOk = body.IndexOf(TailText, StringComparison.Ordinal) >= 0
                    && !SpanContains(body, TailText, "歧义无法操作");
                bool rightGeo = HasGeo(body);
                Ok(run, "wu-cell-cr9", leftOk && rightGeo, Preview(body));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunOverlap(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-overlap-cr"))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, new string('\r', 8))
                    || !s.InsertTable(path, 1, 1, new[] { new string('\r', 9) }))
                {
                    Ok(run, "wu-overlap-cr", false, "COM 写重叠 \\r 失败");
                    return;
                }

                string body = ReadBody(s, ch);
                Ok(run, "wu-overlap-cr",
                    HasGeo(body) || body.IndexOf("歧义无法操作", StringComparison.Ordinal) >= 0,
                    Preview(body));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunVacuum(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-vacuum", "wu-fmt-vacuum"))
            {
                return;
            }

            try
            {
                if (!s.InsertTable(path, 2, 2, new[] { Title, "", "备注", "" }))
                {
                    FailMany(run, "COM 真空表失败", "wu-vacuum", "wu-fmt-vacuum");
                    return;
                }

                string body = ReadBody(s, ch);
                string geo = FirstGeo(body);
                bool vacuum = !string.IsNullOrEmpty(geo)
                    && body.IndexOf("[" + geo + "-start][EMPTY][" + geo + "-end]", StringComparison.Ordinal) >= 0;
                Ok(run, "wu-vacuum", vacuum && body.IndexOf(Title, StringComparison.Ordinal) >= 0, Preview(body));
                ToolResult fmt = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = ch,
                    ["target_codes"] = geo
                });
                Ok(run, "wu-fmt-vacuum",
                    !fmt.Success && Contains(fmt.Error, "empty_cell_no_format"),
                    fmt.Success ? "真空取格式不该成功" : (fmt.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunLongCell(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-long-cell"))
            {
                return;
            }

            try
            {
                string longText = new string('字', 160) + "。";
                if (!s.InsertTable(path, 1, 2, new[] { longText, longText }))
                {
                    Ok(run, "wu-long-cell", false, "COM 长格失败");
                    return;
                }

                string body = ReadBody(s, ch);
                Ok(run, "wu-long-cell",
                    CountGeoStarts(body) >= 2 && body.IndexOf("字字字", StringComparison.Ordinal) >= 0,
                    Preview(body));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunWrites(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-write-merged", "wu-reread"))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, Merge))
                {
                    FailMany(run, "COM 写并段稿失败", "wu-write-merged", "wu-reread");
                    return;
                }

                string body1 = ReadBody(s, ch);
                string body2 = ReadBody(s, ch);
                Ok(run, "wu-reread",
                    body1.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0
                    && body2.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0
                    && CountSStarts(body1) == CountSStarts(body2),
                    Preview(body2));
                string code = FirstS(body1);
                ToolResult rep = Replace(s, ch, code, FillMerged);
                Ok(run, "wu-write-merged",
                    rep.Success && ReadBody(s, ch).IndexOf(FillMerged, StringComparison.Ordinal) >= 0,
                    rep.Success ? "并码写入" : (rep.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunParaCr(TestRun run, WdSession s)
        {
            string path, ch;
            if (!TryOpen(run, s, out path, out ch, "wu-para-cr"))
            {
                return;
            }

            try
            {
                string left = TailText + new string('\r', 8);
                string right = new string('\r', 9);
                if (!s.InsertTable(path, 1, 2, new[] { left, right }))
                {
                    Ok(run, "wu-para-cr", false, "COM 段编码表失败");
                    return;
                }

                ToolResult got = s.GetContent(ch, "paragraph");
                string body = got.Success ? WdSession.Str(got, "document_content") : "";
                Ok(run, "wu-para-cr",
                    got.Success
                    && body.IndexOf("[P_", StringComparison.Ordinal) >= 0
                    && body.IndexOf("歧义无法操作", StringComparison.Ordinal) < 0
                    && Regex.IsMatch(body, @"P_[0-9A-Za-z]+_\d+"),
                    got.Success ? Preview(body) : (got.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static bool WriteDupAroundTable(WdSession s, string path)
        {
            if (s.IsWps)
            {
                return s.TryWpsDoc(path, doc =>
                {
                    object content = GetProp(doc, "Content");
                    SetProp(content, "Text", Dup + "\r" + Dup);
                    object paras = GetProp(doc, "Paragraphs");
                    int n = Convert.ToInt32(GetProp(paras, "Count"));
                    if (n < 2)
                    {
                        Invoke(content, "InsertAfter", Dup);
                        paras = GetProp(doc, "Paragraphs");
                    }

                    object para2 = Invoke(paras, "Item", 2);
                    object at = GetProp(para2, "Range");
                    Invoke(at, "Collapse", 1);
                    object tables = GetProp(doc, "Tables");
                    object table = Invoke(tables, "Add", at, 1, 1);
                    object cell = Invoke(table, "Cell", 1, 1);
                    SetProp(GetProp(cell, "Range"), "Text", Splitter);
                    return StampLateDupOutsideTable(doc);
                });
            }

            return s.OnWordDoc(path, doc =>
            {
                doc.Content.Text = Dup + "\r" + Dup;
                if (doc.Paragraphs.Count < 2)
                {
                    doc.Content.InsertAfter(Dup);
                }

                Word.Range at = doc.Paragraphs[2].Range;
                at.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                Word.Table table = doc.Tables.Add(at, 1, 1);
                table.Cell(1, 1).Range.Text = Splitter;
                return StampWordDupOutsideTable(doc);
            });
        }

        private const int WdWithInTable = 12;

        private static bool StampWordDupOutsideTable(Word.Document doc)
        {
            Word.Paragraph first = null;
            Word.Paragraph last = null;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (ParaInTableWord(p))
                {
                    continue;
                }

                if (first == null)
                {
                    first = p;
                }

                last = p;
            }

            if (first == null)
            {
                return false;
            }

            if (last == first)
            {
                doc.Content.InsertAfter(Dup);
                last = doc.Paragraphs[doc.Paragraphs.Count];
                if (ParaInTableWord(last))
                {
                    return false;
                }
            }

            SetParaTextWord(first, Dup);
            SetParaTextWord(last, Dup);
            return true;
        }

        private static bool StampLateDupOutsideTable(object doc)
        {
            object paras = GetProp(doc, "Paragraphs");
            int n = Convert.ToInt32(GetProp(paras, "Count"));
            object first = null;
            object last = null;
            for (int i = 1; i <= n; i++)
            {
                object p = Invoke(paras, "Item", i);
                if (ParaInTableLate(p))
                {
                    continue;
                }

                if (first == null)
                {
                    first = p;
                }

                last = p;
            }

            if (first == null)
            {
                return false;
            }

            if (object.ReferenceEquals(first, last))
            {
                Invoke(GetProp(doc, "Content"), "InsertAfter", Dup);
                paras = GetProp(doc, "Paragraphs");
                n = Convert.ToInt32(GetProp(paras, "Count"));
                last = Invoke(paras, "Item", n);
                if (ParaInTableLate(last))
                {
                    return false;
                }
            }

            SetParaTextLate(first, Dup);
            SetParaTextLate(last, Dup);
            return true;
        }

        private static bool ParaInTableWord(Word.Paragraph p)
        {
            try
            {
                return (bool)p.Range.Information[Word.WdInformation.wdWithInTable];
            }
            catch
            {
                return p.Range.Tables.Count > 0;
            }
        }

        private static bool ParaInTableLate(object para)
        {
            object range = GetProp(para, "Range");
            try
            {
                object info = range.GetType().InvokeMember(
                    "Information",
                    System.Reflection.BindingFlags.GetProperty,
                    null,
                    range,
                    new object[] { WdWithInTable });
                return Convert.ToBoolean(info);
            }
            catch
            {
                try
                {
                    return Convert.ToInt32(GetProp(GetProp(range, "Tables"), "Count")) > 0;
                }
                catch
                {
                    return false;
                }
            }
        }

        private static void SetParaTextWord(Word.Paragraph p, string text)
        {
            Word.Range r = p.Range.Duplicate;
            if (r.End > r.Start)
            {
                r.End = r.End - 1;
            }

            r.Text = text;
        }

        private static void SetParaTextLate(object para, string text)
        {
            object r = GetProp(para, "Range");
            int start = Convert.ToInt32(GetProp(r, "Start"));
            int end = Convert.ToInt32(GetProp(r, "End"));
            if (end > start)
            {
                Invoke(r, "SetRange", start, end - 1);
            }

            SetProp(r, "Text", text);
        }

        private static ToolResult Replace(WdSession s, string id, string code, string neu)
        {
            return s.ProcessActions(id, new List<Dictionary<string, object>>
            {
                new Dictionary<string, object>
                {
                    ["type"] = "replace",
                    ["detail"] = new Dictionary<string, object>
                    {
                        ["original"] = new Dictionary<string, object> { ["codes"] = code ?? "" },
                        ["new"] = neu,
                        ["format"] = DefaultFormat()
                    }
                }
            });
        }

        private static bool TryOpen(TestRun run, WdSession s, out string path, out string id, params string[] failIds)
        {
            path = s.NewDocxPath();
            id = null;
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                FailMany(run, opened.Error ?? "新建失败", failIds);
                return false;
            }

            id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            return true;
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static Dictionary<string, object> CreateArgs(string channelId, string xmlPath)
        {
            var args = new Dictionary<string, object> { ["path"] = xmlPath ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static string TwoCellXml(string a, string b)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wu-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
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

        private static bool HasS(string body)
        {
            return body.IndexOf("[S_", StringComparison.Ordinal) >= 0;
        }

        private static bool HasGeo(string body)
        {
            return Regex.IsMatch(body ?? "", @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
        }

        private static int CountSStarts(string body)
        {
            return Regex.Matches(body ?? "", @"\[S_[0-9A-Za-z]+-start").Count;
        }

        private static string FirstS(string body)
        {
            Match m = Regex.Match(body ?? "", @"\[(S_[0-9A-Za-z]+)-start");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string D22Containing(string body, string needle)
        {
            foreach (Match m in Regex.Matches(
                body ?? "",
                @"\[(S_[0-9A-Za-z]+)-start 歧义无法操作\](.*?)\[S_[0-9A-Za-z]+-end\]",
                RegexOptions.Singleline))
            {
                if (m.Groups[2].Value.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return m.Groups[1].Value;
                }
            }

            Match any = Regex.Match(body ?? "", @"\[(S_[0-9A-Za-z]+)-start 歧义无法操作\]");
            return any.Success ? any.Groups[1].Value : "";
        }

        private static int CountGeoStarts(string body)
        {
            return Regex.Matches(body ?? "", @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-start").Count;
        }

        private static string Err(ToolResult r)
        {
            if (r == null)
            {
                return "null";
            }

            string e = r.Error ?? "";
            if (r.Data != null)
            {
                try
                {
                    e = e + " " + JsonConvert.SerializeObject(r.Data);
                }
                catch
                {
                    e = e + " " + Convert.ToString(r.Data);
                }
            }

            return e.Trim();
        }

        private static string FirstGeo(string body)
        {
            Match m = Regex.Match(body ?? "", @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
            return m.Success ? m.Value : "";
        }

        private static bool SpanContains(string display, string needle, string extra)
        {
            int hit = display.IndexOf(needle, StringComparison.Ordinal);
            if (hit < 0)
            {
                return false;
            }

            int start = display.LastIndexOf("-start]", hit, StringComparison.Ordinal);
            int end = display.IndexOf("-end]", hit, StringComparison.Ordinal);
            if (start < 0 || end < 0 || end < start)
            {
                return false;
            }

            string inner = display.Substring(start, end - start);
            return inner.IndexOf(extra, StringComparison.Ordinal) >= 0;
        }

        private static string Preview(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return "空 display";
            }

            return body.Length <= 220 ? body : body.Substring(0, 220);
        }

        private static string TrimBlob(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            return text.Length <= 280 ? text : text.Substring(0, 280);
        }

        private static string QuoteCr(string text)
        {
            if (text == null)
            {
                return "(null)";
            }

            return "\"" + text.Replace("\r", "\\r").Replace("\n", "\\n").Replace("\a", "\\a") + "\"";
        }

        private static void DumpTdBodies(string display)
        {
            int i = 0;
            foreach (Match td in Regex.Matches(
                display ?? "",
                @"<t[dh][^>]*>(.*?)</t[dh]>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                i++;
                Console.WriteLine("  td[" + i + "] " + QuoteCr(td.Groups[1].Value));
            }
        }

        private static void DumpParagraphBlocks(string display)
        {
            foreach (Match m in Regex.Matches(
                display ?? "",
                @"\[(P_[0-9A-Za-z_]+) start\](.*?)\[\1 end\]",
                RegexOptions.Singleline))
            {
                string inner = Regex.Replace(m.Groups[2].Value, @"\[in_table=[^\]]+\]", "");
                Console.WriteLine("  P_ " + m.Groups[1].Value + " inner=" + QuoteCr(inner));
            }
        }

        private static List<string> CollectEmptyCellParagraphCodes(string display)
        {
            var codes = new List<string>();
            if (string.IsNullOrEmpty(display))
            {
                return codes;
            }

            foreach (Match td in Regex.Matches(
                display,
                @"<t[dh][^>]*>(.*?)</t[dh]>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                string body = td.Groups[1].Value ?? "";
                string stripped = Regex.Replace(body, @"\[P_[0-9A-Za-z_]+ (?:start|end)\]", "");
                stripped = Regex.Replace(stripped, @"\[in_table=[^\]]+\]", "");
                stripped = Regex.Replace(stripped, @"\[T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+-(?:start|end)\]", "");
                stripped = Regex.Replace(stripped, @"[\r\n\a\t ]+", "");
                if (stripped.Length > 0)
                {
                    continue;
                }

                foreach (Match p in Regex.Matches(body, @"\[(P_[0-9A-Za-z_]+) start\]"))
                {
                    if (!codes.Contains(p.Groups[1].Value))
                    {
                        codes.Add(p.Groups[1].Value);
                    }
                }
            }

            return codes;
        }

        private static bool Contains(string text, string needle)
        {
            return !string.IsNullOrEmpty(text)
                && text.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void Ok(TestRun run, string id, bool cond, string reason)
        {
            if (cond)
            {
                run.CaseOk(id);
            }
            else
            {
                run.CaseFail(id, reason ?? "");
            }
        }

        private static void FailMany(TestRun run, string reason, params string[] ids)
        {
            foreach (string id in ids)
            {
                run.CaseFail(id, reason);
            }
        }

        private static object GetProp(object target, string name)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty,
                null,
                target,
                null);
        }

        private static void SetProp(object target, string name, object value)
        {
            target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.SetProperty,
                null,
                target,
                new[] { value });
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.InvokeMethod,
                null,
                target,
                args);
        }
    }
}
