using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1;
using Word = Microsoft.Office.Interop.Word;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 6 波：XML 建表、填空槽、抽/套格式、删表重建、新建后第一次 optimize。
    /// </summary>
    internal static class WtCases
    {
        public const int Target = 147;

        private const string LabelA = "易写表指标";
        private const string LabelB = "易写表WACC";
        private const string FillA = "8.46%";
        private const string FillB = "12.00%";
        private const string MergeHead = "易写表合并头";
        private const string MergeLeft = "易写表左";
        private const string MergeRight = "易写表右";
        private const string AfterSent = "易写表前句。";
        private const string AfterSentB = "易写表后句。";
        private const string RebuildA = "易写表重建甲";
        private const string RebuildB = "易写表重建乙";
        private const string ConsA = "易写表巩固甲";
        private const string ConsB = "易写表巩固乙";
        private const string ConsC = "易写表巩固丙";
        private const string ConsD = "易写表巩固丁";
        private const string ConsSent = "易写表巩固句。";
        private const string Cell1 = "易写表单格";
        private const string Row3A = "易写表三行A";
        private const string Row3B = "易写表三行B";
        private const string Row3C = "易写表三行C";
        private const string SpanTxt = "易写表重点";
        private const string RsLeft = "易写表跨行左";
        private const string RsUp = "易写表跨行上";
        private const string RsDown = "易写表跨行下";
        private const string ConsE = "易写表巩固戊";
        private const string MoreStart = "易写表文首格";
        private const string WideA = "易写表宽甲";
        private const string WideB = "易写表宽乙";
        private const string WideC = "易写表宽丙";
        private const string TallA = "易写表高甲";
        private const string TallB = "易写表高乙";
        private const string PairA = "易写表对甲";
        private const string PairB = "易写表对乙";

        private static readonly string[] ExtraIds =
        {
            "wt-create-filename", "wt-create-root", "wt-create-colw-gen", "wt-create-colw-tbl",
            "wt-create-style-prop", "wt-create-pos-end", "wt-create-unk-code", "wt-create-unk-tbl",
            "wt-create-no-data", "wt-create-1x1", "wt-create-3row", "wt-create-span",
            "wt-create-rowspan", "wt-create-two", "wt-create-two-text", "wt-create-after-tbl",
            "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
            "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-unk-t", "wt-read-unk-t",
            "wt-read-slots", "wt-read-rows-n", "wt-fill-blank", "wt-fill-row2",
            "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty", "wt-fill-del",
            "wt-xfmt-no-sel", "wt-xfmt-ctx", "wt-xfmt-unk", "wt-xfmt-no-type",
            "wt-afmt-no-id", "wt-afmt-unk", "wt-afmt-bad-xml", "wt-afmt-path",
            "wt-afmt-applied", "wt-afmt-echo", "wt-del-unk", "wt-del-twice",
            "wt-opt-unk", "wt-opt-after-del", "wt-opt-rebuild", "wt-del-keep-sent",
            "wt-read-gone", "wt-two-del-first"
        };

        private static readonly string[] ExtraIds2 =
        {
            "wt-solo-after", "wt-solo-after-id", "wt-solo-fill", "wt-solo-fill-text", "wt-solo-del-keep",
            "wt-two-after-first", "wt-two-keep-both", "wt-two-order", "wt-two-after-last", "wt-two-last-order",
            "wt-before-tbl", "wt-before-tbl-id", "wt-before-keep-sent", "wt-before-order",
            "wt-read-first-after-before", "wt-third-after-tbl", "wt-third-text",
            "wt-cursor-after-two", "wt-cursor-first-text", "wt-cursor-n", "wt-no-join-labels",
            "wt-fill-first-late", "wt-fill-first-late-text", "wt-read-second", "wt-fill-second",
            "wt-fill-second-text", "wt-opt-second", "wt-opt-first-late", "wt-xfmt-second",
            "wt-afmt-first-late", "wt-afmt-second-keep", "wt-del-second-keep-first",
            "wt-del-first-keep-second", "wt-start-keep-old",
            "wt-create-unk-chart", "wt-create-unk-image", "wt-create-two-modes",
            "wt-create-empty-codes", "wt-create-empty-table",
            "wt-create-2x3", "wt-create-3x2", "wt-create-2x1", "wt-read-2x3-rows",
            "wt-fill-2x3", "wt-opt-2x3", "wt-del-2x3-gone",
            "wt-create-bad-merge", "wt-create-empty-tr", "wt-fill-second-short",
            "wt-create-after-filled",
            "wt-nested-create-id", "wt-nested-after-ins", "wt-nested-after-in-cell"
        };

        public static void RunEmptyCell(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            const string Title = "实验操作方法和步骤";
            const string Fill = "1. 数据审核与注册：";
            const string TitleInsert = "（续写）";
            var ids = new[]
            {
                "wt-ec-display", "wt-ec-replace", "wt-ec-insert", "wt-ec-title-insert",
                "wt-ec-format-fail", "wt-ec-mix-fail", "wt-ec-no-format", "wt-ec-after",
                "wt-ec-table-id"
            };
            var s = new WdSession(host);
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            try
            {
                s.ActivateDoc(path);
                ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(Title, "", "备注", ""))));
                string tableId = WdSession.Str(created, "table_id");
                if (!created.Success || string.IsNullOrEmpty(tableId))
                {
                    FailMany(run, created.Error ?? "建表失败", ids);
                    return;
                }

                string body = ReadBody(s, id);
                Match emptyMatch = Regex.Match(body, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
                string emptyCode = emptyMatch.Success ? emptyMatch.Value : "";
                Ok(run, "wt-ec-display",
                    body.IndexOf("[S_", StringComparison.Ordinal) >= 0
                        && emptyMatch.Success
                        && body.IndexOf("[" + emptyCode + "-start][EMPTY][" + emptyCode + "-end]", StringComparison.Ordinal) >= 0,
                    emptyMatch.Success ? emptyCode : "display 无空格码");

                Match titleMatch = Regex.Match(body, @"\[(S_[0-9a-zA-Z]{5})-start\]" + Regex.Escape(Title));
                string titleCode = titleMatch.Success ? titleMatch.Groups[1].Value : "";

                ToolResult fmt = s.GetFormatContext(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["target_codes"] = emptyCode
                });
                Ok(run, "wt-ec-format-fail",
                    !fmt.Success && Contains(fmt.Error, "empty_cell_no_format"),
                    fmt.Success ? "取格式不该成功" : (fmt.Error ?? ""));

                if (!string.IsNullOrEmpty(titleCode) && !string.IsNullOrEmpty(emptyCode))
                {
                    ToolResult mix = s.ProcessActions(id, new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            ["type"] = "replace",
                            ["detail"] = new Dictionary<string, object>
                            {
                                ["original"] = new Dictionary<string, object>
                                {
                                    ["codes"] = titleCode + "," + emptyCode
                                },
                                ["new"] = Fill,
                                ["format"] = DefaultFormat()
                            }
                        }
                    });
                    Ok(run, "wt-ec-mix-fail",
                        !mix.Success && StillHas(s, id, Title) && !StillHas(s, id, Fill),
                        mix.Success ? "混码不该写入" : (mix.Error ?? ""));
                }
                else
                {
                    Ok(run, "wt-ec-mix-fail", false, "缺 title/empty 码");
                }

                ToolResult noFmt = s.ProcessActions(id, new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "replace",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["original"] = new Dictionary<string, object> { ["codes"] = emptyCode },
                            ["new"] = Fill
                        }
                    }
                });
                Ok(run, "wt-ec-no-format",
                    !noFmt.Success && !StillHas(s, id, Fill),
                    noFmt.Success ? "缺 format 不该写入" : (noFmt.Error ?? ""));

                ToolResult afterTbl = s.ProcessActions(id, new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "insert",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["target"] = new Dictionary<string, object> { ["table"] = tableId },
                            ["insert"] = AfterSent,
                            ["format"] = DefaultFormat()
                        }
                    }
                });
                Ok(run, "wt-ec-table-id",
                    afterTbl.Success && StillHas(s, id, AfterSent),
                    afterTbl.Success ? "表后插入" : (afterTbl.Error ?? ""));

                if (!string.IsNullOrEmpty(titleCode))
                {
                    ToolResult titleIns = s.ProcessActions(id, new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            ["type"] = "insert",
                            ["detail"] = new Dictionary<string, object>
                            {
                                ["target"] = new Dictionary<string, object> { ["codes"] = titleCode },
                                ["insert"] = TitleInsert,
                                ["format"] = DefaultFormat()
                            }
                        }
                    });
                    string afterTitle = ReadBody(s, id);
                    Ok(run, "wt-ec-title-insert",
                        titleIns.Success
                            && afterTitle.IndexOf(Title + TitleInsert, StringComparison.Ordinal) >= 0
                            && afterTitle.IndexOf(emptyCode, StringComparison.Ordinal) >= 0,
                        titleIns.Success ? "题后应在左格且右格仍有空码" : (titleIns.Error ?? ""));
                }
                else
                {
                    Ok(run, "wt-ec-title-insert", false, "无题目 S_");
                }

                ToolResult ins = s.ProcessActions(id, new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "insert",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["target"] = new Dictionary<string, object> { ["codes"] = emptyCode },
                            ["insert"] = Fill,
                            ["format"] = DefaultFormat()
                        }
                    }
                });
                string afterIns = ReadBody(s, id);
                Ok(run, "wt-ec-insert",
                    ins.Success
                        && afterIns.IndexOf(Fill, StringComparison.Ordinal) >= 0
                        && afterIns.IndexOf(Title + Fill, StringComparison.Ordinal) < 0,
                    ins.Success ? "insert 应填右格" : (ins.Error ?? ""));

                string afterInsBody = ReadBody(s, id);
                Ok(run, "wt-ec-after",
                    afterInsBody.IndexOf("[" + emptyCode + "-start]", StringComparison.Ordinal) < 0
                        && afterInsBody.IndexOf("[S_", StringComparison.Ordinal) >= 0
                        && afterInsBody.IndexOf(Fill, StringComparison.Ordinal) >= 0,
                    "填后空格码应消失并出现 S_");

                string path2;
                string id2;
                if (TryOpen(run, s, out path2, out id2, "wt-ec-replace"))
                {
                    s.ActivateDoc(path2);
                    s.CreateTableFromXml(CreateArgs(id2, WriteXml(GridXml(Title, "", "备注", ""))));
                    string body2 = ReadBody(s, id2);
                    Match empty2 = Regex.Match(body2, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
                    string code2 = empty2.Success ? empty2.Value : "";
                    ToolResult rep = s.ProcessActions(id2, new List<Dictionary<string, object>>
                    {
                        new Dictionary<string, object>
                        {
                            ["type"] = "replace",
                            ["detail"] = new Dictionary<string, object>
                            {
                                ["original"] = new Dictionary<string, object> { ["codes"] = code2 },
                                ["new"] = Fill,
                                ["format"] = DefaultFormat()
                            }
                        }
                    });
                    string afterRep = ReadBody(s, id2);
                    Ok(run, "wt-ec-replace",
                        rep.Success
                            && afterRep.IndexOf(Fill, StringComparison.Ordinal) >= 0
                            && afterRep.IndexOf(Title + Fill, StringComparison.Ordinal) < 0,
                        rep.Success ? "replace 应填右格" : (rep.Error ?? ""));
                    s.CloseDocByPath(path2);
                }

                string wordDetail, wpsDetail;
                bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
                bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
                Ok(run, "wt-ec-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        public static void RunGeoPlan(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            const string Title = "实验操作方法和步骤";
            const string CellA = "格甲内容。";
            const string CellB = "格乙内容。";
            const string TitleNew = "实验操作方法和步骤改。";
            var ids = new[]
            {
                "wt-plan-two", "wt-plan-split", "wt-plan-delete", "wt-plan-no-quit"
            };
            var s = new WdSession(host);
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            try
            {
                s.ActivateDoc(path);
                ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(Title, "", "备注", ""))));
                string tableId = WdSession.Str(created, "table_id");
                if (!created.Success || string.IsNullOrEmpty(tableId))
                {
                    FailMany(run, created.Error ?? "建表失败", ids);
                    return;
                }

                string body = ReadBody(s, id);
                MatchCollection empties = Regex.Matches(body, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
                string codeA = empties.Count > 0 ? empties[0].Value : "";
                string codeB = empties.Count > 1 ? empties[1].Value : "";
                Match titleMatch = Regex.Match(body, @"\[(S_[0-9a-zA-Z]{5})-start\]" + Regex.Escape(Title));
                string titleCode = titleMatch.Success ? titleMatch.Groups[1].Value : "";

                if (string.IsNullOrEmpty(codeA) || string.IsNullOrEmpty(codeB))
                {
                    FailMany(run, "缺两个空格码", "wt-plan-two", "wt-plan-split");
                }
                else
                {
                    ToolResult two = s.ProcessActions(id, new List<Dictionary<string, object>>
                    {
                        GeoReplace(codeA, CellA),
                        GeoReplace(codeB, CellB)
                    });
                    string afterTwo = ReadBody(s, id);
                    Ok(run, "wt-plan-two",
                        two.Success
                            && afterTwo.IndexOf(CellA, StringComparison.Ordinal) >= 0
                            && afterTwo.IndexOf(CellB, StringComparison.Ordinal) >= 0,
                        two.Success ? "同表两格一次调用" : (two.Error ?? ""));
                }

                string pathSplit, idSplit;
                if (!TryOpen(run, s, out pathSplit, out idSplit, "wt-plan-split"))
                {
                    return;
                }

                try
                {
                    s.ActivateDoc(pathSplit);
                    s.CreateTableFromXml(CreateArgs(idSplit, WriteXml(GridXml(Title, "", "备注", ""))));
                    string splitBody = ReadBody(s, idSplit);
                    MatchCollection splitEmpties = Regex.Matches(splitBody, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
                    string splitA = splitEmpties.Count > 0 ? splitEmpties[0].Value : "";
                    string splitB = splitEmpties.Count > 1 ? splitEmpties[1].Value : "";
                    Match splitTitle = Regex.Match(splitBody, @"\[(S_[0-9a-zA-Z]{5})-start\]" + Regex.Escape(Title));
                    string splitTitleCode = splitTitle.Success ? splitTitle.Groups[1].Value : "";
                    if (string.IsNullOrEmpty(splitA) || string.IsNullOrEmpty(splitB) || string.IsNullOrEmpty(splitTitleCode))
                    {
                        Ok(run, "wt-plan-split", false, "缺码");
                    }
                    else
                    {
                        ToolResult split = s.ProcessActions(idSplit, new List<Dictionary<string, object>>
                        {
                            GeoReplace(splitA, CellA),
                            new Dictionary<string, object>
                            {
                                ["type"] = "replace",
                                ["detail"] = new Dictionary<string, object>
                                {
                                    ["original"] = new Dictionary<string, object> { ["codes"] = splitTitleCode },
                                    ["new"] = TitleNew,
                                    ["format"] = DefaultFormat()
                                }
                            },
                            GeoReplace(splitB, CellB)
                        });
                        string afterSplit = ReadBody(s, idSplit);
                        Ok(run, "wt-plan-split",
                            split.Success
                                && afterSplit.IndexOf(CellA, StringComparison.Ordinal) >= 0
                                && afterSplit.IndexOf(CellB, StringComparison.Ordinal) >= 0
                                && afterSplit.IndexOf(TitleNew, StringComparison.Ordinal) >= 0,
                            split.Success ? "夹句仍两格+题都写上" : (split.Error ?? ""));
                    }
                }
                finally
                {
                    s.CloseDocByPath(pathSplit);
                }

                string pathDel, idDel;
                if (!TryOpen(run, s, out pathDel, out idDel, "wt-plan-delete"))
                {
                    return;
                }

                try
                {
                    s.ActivateDoc(pathDel);
                    s.CreateTableFromXml(CreateArgs(idDel, WriteXml(GridXml("同", "同", "备", ""))));
                    string delBody = ReadBody(s, idDel);
                    Match delGeo = Regex.Match(delBody, @"T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+");
                    string delCode = delGeo.Success ? delGeo.Value : "";
                    if (string.IsNullOrEmpty(delCode))
                    {
                        Ok(run, "wt-plan-delete", false, "无几何码");
                    }
                    else
                    {
                        ToolResult del = s.ProcessActions(idDel, new List<Dictionary<string, object>>
                        {
                            new Dictionary<string, object>
                            {
                                ["type"] = "delete",
                                ["detail"] = new Dictionary<string, object>
                                {
                                    ["deletion"] = new Dictionary<string, object> { ["codes"] = delCode }
                                }
                            }
                        });
                        string afterDel = ReadBody(s, idDel);
                        bool cleared = afterDel.IndexOf("[" + delCode + "-start][EMPTY][" + delCode + "-end]", StringComparison.Ordinal) >= 0
                            || afterDel.IndexOf("[" + delCode + "-start][" + delCode + "-end]", StringComparison.Ordinal) >= 0
                            || (afterDel.Split(new[] { "同" }, StringSplitOptions.None).Length - 1) < (delBody.Split(new[] { "同" }, StringSplitOptions.None).Length - 1);
                        Ok(run, "wt-plan-delete",
                            del.Success && cleared,
                            del.Success ? "delete 几何码应清格" : (del.Error ?? ""));
                    }
                }
                finally
                {
                    s.CloseDocByPath(pathDel);
                }

                string wordDetail, wpsDetail;
                bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
                bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
                Ok(run, "wt-plan-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static Dictionary<string, object> GeoReplace(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["codes"] = code },
                    ["new"] = text,
                    ["format"] = DefaultFormat()
                }
            };
        }

        public static void RunCreateAfterOnly(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warmPath = WarmWpsChannel(s);
            try
            {
                string path, id;
                if (!TryOpen(run, s, out path, out id, "wt-create-after"))
                {
                    return;
                }

                if (!s.TypePlain(path, AfterSent))
                {
                    FailMany(run, "COM 写句失败", "wt-create-after");
                    s.CloseDocByPath(path);
                    return;
                }

                s.ActivateDoc(path);
                ToolResult got = s.GetContent(id);
                string sent = CodeContaining(WdSession.Str(got, "document_content"), AfterSent, "S");
                if (string.IsNullOrEmpty(sent))
                {
                    FailMany(run, "无可抄含前句的 S_", "wt-create-after");
                    s.CloseDocByPath(path);
                    return;
                }

                var afterArgs = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
                afterArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult after = s.CreateTableFromXml(afterArgs);
                Ok(run, "wt-create-after",
                    after.Success && StillHas(s, id, AfterSent) && StillHas(s, id, LabelA),
                    after.Success ? "句后建表丢了句或表" : (after.Error ?? ""));
                s.CloseDocByPath(path);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warmPath))
                {
                    s.CloseDocByPath(warmPath);
                }
            }
        }

        /// <summary>
        /// 两句话：存文 ≠ 全文，必须走原来的 Find，不能误用整篇 Content。
        /// 句后建表应夹在两句之间（第一句后）或落在第二句后。
        /// </summary>
        /// <summary>对照巩固套：首张表刚建完 vs 后面又建几张表之后，读空槽。</summary>
        public static void RunEmptySlotDump(string host)
        {
            var s = new WdSession(host);
            var dummy = new TestRun();
            string path, id;
            if (!TryOpen(dummy, s, out path, out id, "wt-wps-empty-slots"))
            {
                Console.WriteLine("open failed");
                return;
            }

            if (!s.TypePlain(path, ConsSent))
            {
                Console.WriteLine("TypePlain failed");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult created = s.CreateTableFromXml(
                CreateArgsFilename(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Console.WriteLine("first id=" + tableId + " ok=" + created.Success);
            DumpRead(s, id, tableId, "just-created");

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = tableId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            Console.WriteLine("after-tbl ok=" + second.Success + " id=" + WdSession.Str(second, "table_id"));
            DumpRead(s, id, tableId, "after-second-table");

            CreateOkHas(s, id, OneCellXml(Cell1), Cell1);
            CreateOkHas(s, id, ThreeRowXml(), Row3A);
            CreateOkHas(s, id, SpanXml(), SpanTxt);
            CreateOkHas(s, id, RowspanXml(), RsLeft);
            Console.WriteLine("firstId still=" + tableId);
            s.WriteTableFacts(path);
            DumpRead(s, id, tableId, "after-extra-tables");

            s.CloseDocByPath(path);
        }

        private static void DumpRead(WdSession s, string id, string tableId, string tag)
        {
            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            Console.WriteLine(tag + " success=" + read.Success + " err=" + (read.Error ?? ""));
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                Console.WriteLine(tag + " rows=null data=" + (data == null ? "null" : data.ToString(Formatting.None)));
                return;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                Console.WriteLine(tag + " row=" + o["row"] + " empty_count=" + o["empty_count"]
                    + " values=" + (o["values"] == null ? "null" : o["values"].ToString(Formatting.None))
                    + " slots=" + (o["slots"] == null ? "null" : o["slots"].ToString(Formatting.None)));
            }

            int er, ec;
            Console.WriteLine(tag + " TryFirstEmpty=" + TryFirstEmptyRow(read, out er, out ec)
                + " row=" + er + " n=" + ec);
        }

        /// <summary>句后、已有表之前再建表，看会不会并进后面那张。</summary>
        public static void RunBeforeTblJoinProbe(string host)
        {
            var s = new WdSession(host);
            var dummy = new TestRun();
            string path, id;
            if (!TryOpen(dummy, s, out path, out id, "wt-before-tbl-join"))
            {
                Console.WriteLine("open failed");
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                Console.WriteLine("TypePlain failed");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.MoveSelectionToEnd(path);
            ToolResult first = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string firstId = WdSession.Str(first, "table_id");
            Console.WriteLine("first success=" + first.Success + " id=" + firstId);
            DumpTablesOrFacts(s, path, "after-first");
            s.WriteSelectionFact(path);

            ToolResult got = s.GetContent(id);
            string sent = CodeContaining(WdSession.Str(got, "document_content"), AfterSent, "S");
            Console.WriteLine("anchor=" + sent);
            var beforeArgs = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            beforeArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult second = s.CreateTableFromXml(beforeArgs);
            Console.WriteLine("before-tbl success=" + second.Success
                + " id=" + WdSession.Str(second, "table_id")
                + " err=" + (second.Error ?? ""));
            DumpTablesOrFacts(s, path, "after-before-insert");
            string body = ReadBody(s, id);
            Console.WriteLine("has句=" + (body.IndexOf(AfterSent, StringComparison.Ordinal) >= 0)
                + " has甲=" + (body.IndexOf(ConsA, StringComparison.Ordinal) >= 0)
                + " has丙=" + (body.IndexOf(ConsC, StringComparison.Ordinal) >= 0));
            Console.WriteLine("doc=" + path);
            s.CloseDocByPath(path);
        }

        private static void DumpTablesOrFacts(WdSession s, string path, string tag)
        {
            Console.WriteLine(tag);
            if (s.IsWps)
            {
                s.WriteTableFacts(path);
                return;
            }

            DumpTables(s, path, tag);
        }

        /// <summary>表后建表后看 COM 是两张还是并成一张。</summary>
        public static void RunAfterTblJoinProbe(string host)
        {
            var s = new WdSession(host);
            var dummy = new TestRun();
            string path, id;
            if (!TryOpen(dummy, s, out path, out id, "wt-after-tbl-join"))
            {
                Console.WriteLine("open failed");
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                Console.WriteLine("TypePlain failed");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult first = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string firstId = WdSession.Str(first, "table_id");
            Console.WriteLine("first success=" + first.Success + " id=" + firstId + " err=" + (first.Error ?? ""));
            DumpTables(s, path, "after-first");

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = firstId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            string secondId = WdSession.Str(second, "table_id");
            Console.WriteLine("second success=" + second.Success + " id=" + secondId + " err=" + (second.Error ?? ""));
            DumpTables(s, path, "after-second");

            string body = ReadBody(s, id);
            Console.WriteLine("has甲=" + (body.IndexOf(ConsA, StringComparison.Ordinal) >= 0)
                + " has乙=" + (body.IndexOf(ConsB, StringComparison.Ordinal) >= 0)
                + " has丙=" + (body.IndexOf(ConsC, StringComparison.Ordinal) >= 0)
                + " has丁=" + (body.IndexOf(ConsD, StringComparison.Ordinal) >= 0));
            Console.WriteLine("T_ count=" + Regex.Matches(body, @"\[T_[0-9a-zA-Z]{5,}-start\]").Count);
            Console.WriteLine("doc left open: " + path);
        }

        private static void DumpTables(WdSession s, string path, string tag)
        {
            s.OnWordDoc(path, doc =>
            {
                Console.WriteLine(tag + " Tables.Count=" + doc.Tables.Count
                    + " Paras=" + doc.Paragraphs.Count
                    + " Content.End=" + doc.Content.End);
                for (int i = 1; i <= doc.Tables.Count; i++)
                {
                    Word.Table t = doc.Tables[i];
                    string title = "";
                    try { title = t.Title ?? ""; } catch { title = "?"; }
                    string cells = "";
                    try
                    {
                        cells = (t.Cell(1, 1).Range.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a")
                            + " | "
                            + (t.Cell(1, 2).Range.Text ?? "").Replace("\r", "\\r").Replace("\a", "\\a");
                    }
                    catch (Exception ex)
                    {
                        cells = ex.Message;
                    }

                    Console.WriteLine("  T" + i + " rows=" + t.Rows.Count + " cols=" + t.Columns.Count
                        + " rng=" + t.Range.Start + "-" + t.Range.End
                        + " title=" + title
                        + " r1=" + cells);
                }

                if (doc.Tables.Count >= 2)
                {
                    int gap = doc.Tables[2].Range.Start - doc.Tables[1].Range.End;
                    Console.WriteLine("  gap(T2.Start-T1.End)=" + gap);
                }

                return true;
            });
        }

        public static void RunCreateAfterTwoOnly(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunAfterTwo(run, s, "wt-create-after-first", AfterSent, AfterSentB, AfterSent);
                RunAfterTwo(run, s, "wt-create-after-last", AfterSent, AfterSentB, AfterSentB);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warmPath))
                {
                    s.CloseDocByPath(warmPath);
                }
            }
        }

        private static void RunAfterTwo(
            TestRun run,
            WdSession s,
            string caseId,
            string first,
            string second,
            string anchorNeedle)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            if (!s.TypePlain(path, first + "\r" + second))
            {
                FailMany(run, "COM 写两句失败", caseId);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult got = s.GetContent(id);
            string body = WdSession.Str(got, "document_content");
            string sent = CodeContaining(body, anchorNeedle, "S");
            if (string.IsNullOrEmpty(sent))
            {
                FailMany(run, "无可抄含锚句的 S_", caseId);
                s.CloseDocByPath(path);
                return;
            }

            var args = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
            args["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult after = s.CreateTableFromXml(args);
            string afterBody = ReadBody(s, id);
            int iFirst = IndexOf(afterBody, first);
            int iSecond = IndexOf(afterBody, second);
            int iTbl = IndexOf(afterBody, LabelA);
            bool order = string.Equals(anchorNeedle, first, StringComparison.Ordinal)
                ? iFirst >= 0 && iTbl > iFirst && iSecond > iTbl
                : iFirst >= 0 && iSecond > iFirst && iTbl > iSecond;
            string reason;
            if (!after.Success)
            {
                reason = after.Error ?? "建表失败";
            }
            else if (!order)
            {
                reason = string.Equals(anchorNeedle, first, StringComparison.Ordinal)
                    ? "表没夹在两句之间（可能误用了全文 Range）"
                    : "表没落在第二句之后";
            }
            else
            {
                reason = "";
            }

            Ok(run, caseId, after.Success && order, reason);
            s.CloseDocByPath(path);
        }

        private static int IndexOf(string text, string needle)
        {
            return string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle)
                ? -1
                : text.IndexOf(needle, StringComparison.Ordinal);
        }

        public static void Run(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunNegEarly(run, s);
            string warmPath = WarmWpsChannel(s);
            try
            {
                RunCreateAndFill(run, s);
                RunFormatDeleteOpt(run, s);
                RunConsolidate(run, s);
                RunConsolidateMore(run, s);
                RunNestedTableAfter(run, s);
                RunNegLate(run, s);
                RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
            }
            finally
            {
                if (!string.IsNullOrEmpty(warmPath))
                {
                    s.CloseDocByPath(warmPath);
                }
            }
        }

        private static string WarmWpsChannel(WdSession s)
        {
            if (!s.IsWps)
            {
                return null;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return null;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.GetContent(id);
            s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            return path;
        }

        private static void BurnFirstWpsOpen(WdSession s)
        {
            if (!s.IsWps)
            {
                return;
            }

            string path = s.NewDocxPath();
            ToolResult opened = s.Open(path, true, s.App);
            if (!opened.Success)
            {
                return;
            }

            string id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(path);
            s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            s.CloseDocByPath(path);
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            string xml = WriteXml(GridXml(LabelA, "", LabelB, ""));
            ToolResult noCh = s.CreateTableFromXml(CreateArgs(null, xml));
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wt-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wt-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wt-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wt-neg-unknown-id", s.CreateTableFromXml(CreateArgs("w999", xml)), "未知");
            ExpectFail(
                run,
                "wt-neg-create-no-path",
                s.CreateTableFromXml(new Dictionary<string, object>()),
                "path");
            ExpectFail(
                run,
                "wt-neg-create-missing",
                s.CreateTableFromXml(CreateArgs(null, Path.Combine(Path.GetTempPath(), "ew-wt-missing-" + Guid.NewGuid().ToString("N") + ".xml"))),
                "不存在");
            ExpectFail(
                run,
                "wt-neg-create-bad-xml",
                s.CreateTableFromXml(CreateArgs(null, WriteXml("<TableConfig><Table><Properties/></Table><Data><tr><td>x</td></tr></Data></TableConfig>"))),
                "Data");
            ExpectFail(
                run,
                "wt-neg-create-pos",
                s.CreateTableFromXml(new Dictionary<string, object>
                {
                    ["path"] = xml,
                    ["position"] = "end"
                }),
                "position");
            ExpectFail(
                run,
                "wt-neg-del-no-id",
                s.DeleteTable(new Dictionary<string, object>()),
                "编号");
            ExpectFail(
                run,
                "wt-neg-read-no-id",
                s.ReadTableRowValues(new Dictionary<string, object>()),
                "table_id");
            ExpectFail(
                run,
                "wt-neg-apply-no-id",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["format"] = DefaultFormat(),
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "table_id");
            ExpectFail(
                run,
                "wt-neg-apply-no-fmt",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "char_format");
            ExpectFail(
                run,
                "wt-neg-apply-inherit",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat(),
                    ["inherit_from"] = "S_00001",
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object> { ["row"] = 1, ["values"] = new[] { FillA } }
                    }
                }),
                "inherit");
            ExpectFail(
                run,
                "wt-neg-apply-no-writes",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat()
                }),
                "writes");
            ExpectFail(
                run,
                "wt-neg-apply-anchor",
                s.ApplyTableRowValues(new Dictionary<string, object>
                {
                    ["table_id"] = "T_00001",
                    ["format"] = DefaultFormat(),
                    ["writes"] = new object[]
                    {
                        new Dictionary<string, object>
                        {
                            ["row"] = 1,
                            ["values"] = new[] { FillA },
                            ["anchor"] = "指标"
                        }
                    }
                }),
                "anchor");
            ExpectFail(
                run,
                "wt-neg-fmt-no-xml",
                s.ApplyTableFormatFromXml(new Dictionary<string, object> { ["table_id"] = "T_00001" }),
                "path");
            ExpectFail(
                run,
                "wt-neg-opt-no-id",
                s.OptimizeTableDisplay(new Dictionary<string, object>()),
                "编号");
        }

        private static void RunCreateAndFill(TestRun run, WdSession s)
        {
            BurnFirstWpsOpen(s);

            string path, id;
            if (!TryOpen(run, s, out path, out id,
                "wt-create-ok", "wt-create-id", "wt-create-text", "wt-create-fields",
                "wt-read-ok", "wt-read-empty", "wt-fill-ok", "wt-fill-text", "wt-fill-keep",
                "wt-fill-count", "wt-fill-wrong-len"))
            {
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);
            ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Ok(run, "wt-create-ok", created.Success, created.Error ?? "建表失败");
            Ok(run, "wt-create-id",
                created.Success && tableId.StartsWith("T_", StringComparison.Ordinal),
                created.Success ? "table_id=" + tableId : (created.Error ?? ""));
            Ok(run, "wt-create-text",
                created.Success && StillHas(s, id, LabelA) && StillHas(s, id, LabelB),
                created.Success ? "建表正文无标签" : (created.Error ?? ""));
            Ok(run, "wt-create-fields",
                created.Success
                    && WdSession.Bool(created, "table_created")
                    && WdSession.IntOf(created, "rows") == 2
                    && WdSession.IntOf(created, "columns") == 2,
                created.Success ? "回包缺 table_created / 行列" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(tableId))
            {
                FailMany(run, created.Error ?? "无 table_id",
                    "wt-read-ok", "wt-read-empty", "wt-fill-ok", "wt-fill-text", "wt-fill-keep",
                    "wt-fill-count", "wt-fill-wrong-len");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            JObject readObj = DataObj(read);
            Ok(run, "wt-read-ok",
                read.Success
                    && WdSession.Bool(read, "do_not_use_getcontent_table_html")
                    && readObj != null
                    && readObj["rows"] is JArray,
                read.Success ? "回包无 do_not_use / rows" : (read.Error ?? ""));

            int emptyRow;
            int emptyCount;
            bool hasEmpty = TryFirstEmptyRow(read, out emptyRow, out emptyCount);
            Ok(run, "wt-read-empty", hasEmpty && emptyCount > 0,
                read.Success ? "没有 empty_count>0 的行" : (read.Error ?? ""));
            if (!hasEmpty)
            {
                FailMany(run, "无空槽可填",
                    "wt-fill-ok", "wt-fill-text", "wt-fill-keep", "wt-fill-count", "wt-fill-wrong-len");
                s.CloseDocByPath(path);
                return;
            }

            int slotCount;
            TryRowSlotCount(read, emptyRow, out slotCount);
            var tooLong = RepeatSlots(slotCount + 1, "x");

            ToolResult badLen = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, tooLong));
            Ok(run, "wt-fill-wrong-len",
                !badLen.Success && StillHas(s, id, LabelA) && !StillHas(s, id, FillA),
                badLen.Success ? "values 个数不对仍写入" : (badLen.Error ?? "错长度却改了字"));

            string[] values = BuildApplyValues(read, emptyRow, FillA);

            ToolResult fill = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, values));
            Ok(run, "wt-fill-ok", fill.Success, fill.Error ?? "填空失败");
            Ok(run, "wt-fill-text",
                fill.Success && StillHas(s, id, FillA),
                fill.Success ? "空槽未写出" : (fill.Error ?? ""));
            Ok(run, "wt-fill-keep",
                fill.Success && StillHas(s, id, LabelA) && StillHas(s, id, LabelB),
                fill.Success ? "填空改了标签格" : (fill.Error ?? ""));
            Ok(run, "wt-fill-count",
                fill.Success && WdSession.IntOf(fill, "cells_written") >= 1,
                fill.Success ? "cells_written=" + WdSession.IntOf(fill, "cells_written") : (fill.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunFormatDeleteOpt(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id,
                "wt-create-after", "wt-create-start", "wt-create-merge",
                "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                "wt-opt-ok", "wt-opt-fields"))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                FailMany(run, "COM 写句失败",
                    "wt-create-after", "wt-create-start", "wt-create-merge",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult got = s.GetContent(id);
            string sent = CodeContaining(WdSession.Str(got, "document_content"), AfterSent, "S");
            if (string.IsNullOrEmpty(sent))
            {
                FailMany(run, "无可抄含前句的 S_",
                    "wt-create-after", "wt-create-start", "wt-create-merge",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            var afterArgs = CreateArgs(id, WriteXml(GridXml(LabelA, "", LabelB, "")));
            afterArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult after = s.CreateTableFromXml(afterArgs);
            string afterId = WdSession.Str(after, "table_id");
            Ok(run, "wt-create-after",
                after.Success && StillHas(s, id, AfterSent) && StillHas(s, id, LabelA),
                after.Success ? "句后建表丢了句或表" : (after.Error ?? ""));

            string startPath = s.NewDocxPath();
            ToolResult startOpen = s.Open(startPath, true, s.App);
            if (!startOpen.Success)
            {
                FailMany(run, startOpen.Error ?? "再建空稿失败", "wt-create-start");
            }
            else
            {
                string startId = WdSession.Str(startOpen, "channel_id");
                var startArgs = CreateArgs(startId, WriteXml(GridXml(LabelA, "1", LabelB, "2")));
                startArgs["target"] = new Dictionary<string, object> { ["position"] = "start" };
                ToolResult start = s.CreateTableFromXml(startArgs);
                Ok(run, "wt-create-start",
                    start.Success && StillHas(s, startId, LabelA),
                    start.Success ? "文首建表无标签" : (start.Error ?? ""));
                s.CloseDocByPath(startPath);
            }

            string mergePath = s.NewDocxPath();
            ToolResult mergeOpen = s.Open(mergePath, true, s.App);
            if (!mergeOpen.Success)
            {
                FailMany(run, mergeOpen.Error ?? "再建空稿失败", "wt-create-merge");
            }
            else
            {
                string mergeId = WdSession.Str(mergeOpen, "channel_id");
                s.ActivateDoc(mergePath);
                ToolResult merge = s.CreateTableFromXml(CreateArgs(mergeId, WriteXml(MergeXml())));
                Ok(run, "wt-create-merge",
                    merge.Success
                        && StillHas(s, mergeId, MergeHead)
                        && StillHas(s, mergeId, MergeLeft)
                        && StillHas(s, mergeId, MergeRight),
                    merge.Success ? "合并表缺可见格" : (merge.Error ?? ""));
                s.CloseDocByPath(mergePath);
            }

            if (!after.Success || string.IsNullOrEmpty(afterId))
            {
                FailMany(run, after.Error ?? "句后表无 id",
                    "wt-xfmt-ok", "wt-xfmt-file", "wt-afmt-ok", "wt-afmt-text", "wt-afmt-no-data",
                    "wt-del-ok", "wt-del-gone", "wt-rebuild-ok", "wt-rebuild-text",
                    "wt-opt-ok", "wt-opt-fields");
                s.CloseDocByPath(path);
                return;
            }

            string extractPath = Path.Combine(Path.GetTempPath(), "ew-wt-fmt-" + Guid.NewGuid().ToString("N") + ".xml");
            ToolResult extract = s.ExtractTableFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = extractPath,
                ["table_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["table_id"] = afterId
                }
            });
            Ok(run, "wt-xfmt-ok", extract.Success, extract.Error ?? "抽格式失败");
            Ok(run, "wt-xfmt-file",
                extract.Success
                    && (WdSession.Bool(extract, "format_saved")
                        || !string.IsNullOrEmpty(WdSession.Str(extract, "format_content"))),
                extract.Success ? "无 format_saved / format_content" : (extract.Error ?? ""));

            ToolResult applyFmt = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId,
                ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor><Style><TableStyle>ThreeLine</TableStyle><HeaderRows>1</HeaderRows></Style></General></TableConfig>"
            });
            Ok(run, "wt-afmt-ok", applyFmt.Success, applyFmt.Error ?? "套表格式失败");
            Ok(run, "wt-afmt-text",
                applyFmt.Success && StillHas(s, id, LabelA) && StillHas(s, id, AfterSent),
                applyFmt.Success ? "套格式改了字" : (applyFmt.Error ?? ""));

            ToolResult applyData = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId,
                ["xml_content"] = GridXml("不该写入甲", "不该写入乙", "不该写入丙", "不该写入丁")
            });
            Ok(run, "wt-afmt-no-data",
                applyData.Success
                    && StillHas(s, id, LabelA)
                    && !StillHas(s, id, "不该写入甲"),
                applyData.Success ? "套格式吃了 Data" : (applyData.Error ?? ""));

            ToolResult opt = s.OptimizeTableDisplay(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId
            });
            Ok(run, "wt-opt-ok", opt.Success, opt.Error ?? "新建后第一次 optimize 失败");
            Ok(run, "wt-opt-fields",
                opt.Success && WdSession.HasKey(opt, "final_font_size"),
                opt.Success ? "回包无 final_font_size" : (opt.Error ?? ""));

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = afterId
            });
            Ok(run, "wt-del-ok", del.Success, del.Error ?? "删表失败");
            Ok(run, "wt-del-gone",
                del.Success && !StillHas(s, id, LabelA),
                del.Success ? "删表后标签还在" : (del.Error ?? ""));

            ToolResult rebuild = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(RebuildA, "1", RebuildB, "2"))));
            Ok(run, "wt-rebuild-ok", rebuild.Success, rebuild.Error ?? "删后重建失败");
            Ok(run, "wt-rebuild-text",
                rebuild.Success && StillHas(s, id, RebuildA) && StillHas(s, id, AfterSent),
                rebuild.Success ? "重建表或前句丢了" : (rebuild.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunConsolidate(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ExtraIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConsSent))
            {
                FailMany(run, "COM 写巩固句失败", ExtraIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);

            ExpectFail(run, "wt-create-root",
                s.CreateTableFromXml(CreateArgs(id, WriteXml("<NotTable><Table><Data><tr><td>x</td></tr></Data></Table></NotTable>"))),
                "TableConfig");
            ExpectFail(run, "wt-create-colw-gen",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General><ColWidths><float>120</float></ColWidths></General><Table><Properties/><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "ColWidths");
            ExpectFail(run, "wt-create-colw-tbl",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><ColWidths><float>120</float></ColWidths><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "ColWidths");
            ExpectFail(run, "wt-create-style-prop",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><Properties><Style><TableStyle>ThreeLine</TableStyle></Style></Properties><Data><tr><td>x</td></tr></Data></Table></TableConfig>"))),
                "Style");
            ExpectFail(run, "wt-create-no-data",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><Properties><ColWidths><float>120</float></ColWidths></Properties></Table></TableConfig>"))),
                "Data");

            var endArgs = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            endArgs["target"] = new Dictionary<string, object> { ["position"] = "end" };
            ExpectFail(run, "wt-create-pos-end", s.CreateTableFromXml(endArgs), "start");

            var unkCode = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkCode["target"] = new Dictionary<string, object> { ["codes"] = "S_99999" };
            ExpectFail(run, "wt-create-unk-code", s.CreateTableFromXml(unkCode), "locate");

            var unkTbl = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkTbl["target"] = new Dictionary<string, object> { ["table"] = "T_99999" };
            ExpectFail(run, "wt-create-unk-tbl", s.CreateTableFromXml(unkTbl), "表格");

            ToolResult created = s.CreateTableFromXml(
                CreateArgsFilename(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string tableId = WdSession.Str(created, "table_id");
            Ok(run, "wt-create-filename",
                created.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsB),
                created.Success ? "filename 建表无标签" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(tableId))
            {
                FailMany(run, created.Error ?? "filename 建表无 table_id",
                    "wt-create-1x1", "wt-create-3row", "wt-create-span", "wt-create-rowspan",
                    "wt-create-two", "wt-create-two-text", "wt-create-after-tbl",
                    "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
                    "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-unk-t", "wt-read-unk-t",
                    "wt-read-slots", "wt-read-rows-n", "wt-fill-blank", "wt-fill-row2",
                    "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty", "wt-fill-del",
                    "wt-xfmt-no-sel", "wt-xfmt-ctx", "wt-xfmt-unk", "wt-xfmt-no-type",
                    "wt-afmt-no-id", "wt-afmt-unk", "wt-afmt-bad-xml", "wt-afmt-path",
                    "wt-afmt-applied", "wt-afmt-echo", "wt-del-unk", "wt-del-twice",
                    "wt-opt-unk", "wt-opt-after-del", "wt-opt-rebuild", "wt-del-keep-sent",
                    "wt-read-gone", "wt-two-del-first");
                s.CloseDocByPath(path);
                return;
            }

            var afterTbl = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            afterTbl["target"] = new Dictionary<string, object> { ["table"] = tableId };
            ToolResult second = s.CreateTableFromXml(afterTbl);
            string secondId = WdSession.Str(second, "table_id");
            Ok(run, "wt-create-after-tbl",
                second.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                second.Success ? "表后建表丢了表" : (second.Error ?? ""));
            Ok(run, "wt-create-two",
                second.Success && secondId.StartsWith("T_", StringComparison.Ordinal) && secondId != tableId,
                second.Success ? "第二张 table_id 不对" : (second.Error ?? ""));
            Ok(run, "wt-create-two-text",
                second.Success && StillHas(s, id, ConsC) && StillHas(s, id, ConsD) && StillHas(s, id, ConsA),
                second.Success ? "两张表标签不全" : (second.Error ?? ""));

            Ok(run, "wt-create-1x1",
                CreateOkHas(s, id, OneCellXml(Cell1), Cell1),
                "单格未出现");
            Ok(run, "wt-create-3row",
                CreateOkHas(s, id, ThreeRowXml(), Row3A) && StillHas(s, id, Row3B) && StillHas(s, id, Row3C),
                "三行标签不全");
            Ok(run, "wt-create-span",
                CreateOkHas(s, id, SpanXml(), SpanTxt),
                "span 格字未出现");
            Ok(run, "wt-create-rowspan",
                CreateOkHas(s, id, RowspanXml(), RsLeft) && StillHas(s, id, RsUp) && StillHas(s, id, RsDown),
                "跨行合并可见格不全");

            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            JObject readObj = DataObj(read);
            JArray rows = readObj == null ? null : readObj["rows"] as JArray;
            Ok(run, "wt-read-slots",
                read.Success && rows != null && RowHasSlots(rows),
                read.Success ? "rows 无 slots" : (read.Error ?? ""));
            Ok(run, "wt-read-rows-n",
                read.Success && rows != null && rows.Count == 2,
                read.Success ? "rows.Count=" + (rows == null ? -1 : rows.Count) : (read.Error ?? ""));

            int emptyRow;
            int emptyCount;
            bool hasEmpty = TryFirstEmptyRow(read, out emptyRow, out emptyCount);
            int row2;
            int empty2;
            bool hasRow2 = TryLastEmptyRow(read, out row2, out empty2);

            ExpectFail(run, "wt-fill-unk-t",
                s.ApplyTableRowValues(FillArgs(id, "T_99999", 1, new[] { "x" })),
                "找不到");
            ExpectFail(run, "wt-read-unk-t",
                s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = "T_99999"
                }),
                "找不到");

            if (!hasEmpty)
            {
                FailMany(run, "巩固表无空槽",
                    "wt-fill-short-len", "wt-fill-index", "wt-fill-bad-row", "wt-fill-row0",
                    "wt-fill-writes-empty", "wt-fill-writes-obj", "wt-fill-blank", "wt-fill-row2",
                    "wt-fill-row2-text", "wt-read-after-fill", "wt-fill-no-empty");
            }
            else
            {
                int slotN;
                TryRowSlotCount(read, emptyRow, out slotN);
                var tooShort = RepeatSlots(Math.Max(0, slotN - 1), "x");

                ToolResult shortLen = s.ApplyTableRowValues(FillArgs(id, tableId, emptyRow, tooShort));
                Ok(run, "wt-fill-short-len",
                    !shortLen.Success && StillHas(s, id, ConsA) && !StillHas(s, id, FillA),
                    shortLen.Success ? "values 偏短仍写入" : (shortLen.Error ?? "错长度却改了字"));

                var indexArgs = FillArgs(id, tableId, emptyRow, BuildApplyValues(read, emptyRow, FillA));
                ((Dictionary<string, object>)((object[])indexArgs["writes"])[0])["index"] = 1;
                ExpectFail(run, "wt-fill-index", s.ApplyTableRowValues(indexArgs), "index");
                ExpectFail(run, "wt-fill-bad-row",
                    s.ApplyTableRowValues(FillArgs(id, tableId, 99, new[] { "x" })),
                    "invalid_row");
                ExpectFail(run, "wt-fill-row0",
                    s.ApplyTableRowValues(FillArgs(id, tableId, 0, new[] { "x" })),
                    "invalid_row");
                ExpectFail(run, "wt-fill-writes-empty",
                    s.ApplyTableRowValues(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["table_id"] = tableId,
                        ["format"] = DefaultFormat(),
                        ["writes"] = new object[0]
                    }),
                    "writes");
                ExpectFail(run, "wt-fill-writes-obj",
                    s.ApplyTableRowValues(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["table_id"] = tableId,
                        ["format"] = DefaultFormat(),
                        ["writes"] = "nope"
                    }),
                    "数组");

                ToolResult blank = s.ApplyTableRowValues(
                    FillArgs(id, tableId, emptyRow, BuildApplyValues(read, emptyRow, "")));
                Ok(run, "wt-fill-blank",
                    blank.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsB) && !StillHas(s, id, FillA),
                    blank.Success ? "全空写入改了标签或写出了数" : (blank.Error ?? ""));

                int fillRow = hasRow2 ? row2 : emptyRow;
                ToolResult fill2 = s.ApplyTableRowValues(
                    FillArgs(id, tableId, fillRow, BuildApplyValues(read, fillRow, FillB)));
                Ok(run, "wt-fill-row2", fill2.Success, fill2.Error ?? "第二空行填空失败");
                Ok(run, "wt-fill-row2-text",
                    fill2.Success && StillHas(s, id, FillB) && !StillHas(s, id, FillA) && StillHas(s, id, ConsA),
                    fill2.Success ? "第二空行字不对" : (fill2.Error ?? ""));

                ToolResult readAfter = s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId
                });
                int afterN;
                TryRowEmptyCount(readAfter, fillRow, out afterN);
                Ok(run, "wt-read-after-fill",
                    readAfter.Success && afterN == 0,
                    readAfter.Success ? "填后 empty_count=" + afterN : (readAfter.Error ?? ""));
                ExpectFail(run, "wt-fill-no-empty",
                    s.ApplyTableRowValues(FillArgs(id, tableId, fillRow, new[] { "x" })),
                    "values_length");
            }

            ExpectFail(run, "wt-xfmt-no-sel",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml")
                }),
                "table_selection");
            ExpectFail(run, "wt-xfmt-ctx",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_context",
                        ["table_id"] = tableId
                    }
                }),
                "by_context");
            ExpectFail(run, "wt-xfmt-unk",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_id",
                        ["table_id"] = "T_99999"
                    }
                }),
                "找不到");
            ExpectFail(run, "wt-xfmt-no-type",
                s.ExtractTableFormat(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wt-x-" + Guid.NewGuid().ToString("N") + ".xml"),
                    ["table_selection"] = new Dictionary<string, object> { ["table_id"] = tableId }
                }),
                "类型");

            ExpectFail(run, "wt-afmt-no-id",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor></General></TableConfig>"
                }),
                "编号");
            ExpectFail(run, "wt-afmt-unk",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = "T_99999",
                    ["xml_content"] = "<TableConfig><General><TableFontColor>000000</TableFontColor></General></TableConfig>"
                }),
                "找不到");
            ExpectFail(run, "wt-afmt-bad-xml",
                s.ApplyTableFormatFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId,
                    ["xml_content"] = "<NotTable/>"
                }),
                "TableConfig");

            string fmtPath = WriteXml(
                "<TableConfig><General><TableFontColor>000000</TableFontColor><Style><TableStyle>ThreeLine</TableStyle><HeaderRows>1</HeaderRows></Style></General></TableConfig>");
            ToolResult afmt = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId,
                ["path"] = fmtPath
            });
            Ok(run, "wt-afmt-path",
                afmt.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsSent),
                afmt.Success ? "path 套格式改了字" : (afmt.Error ?? ""));
            Ok(run, "wt-afmt-applied",
                afmt.Success && WdSession.Bool(afmt, "table_format_applied"),
                afmt.Success ? "无 table_format_applied" : (afmt.Error ?? ""));
            Ok(run, "wt-afmt-echo",
                afmt.Success && Contains(WdSession.Str(afmt, "table_id"), tableId),
                afmt.Success ? "回包 table_id 不对" : (afmt.Error ?? ""));

            ExpectFail(run, "wt-del-unk",
                s.DeleteTable(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = "T_99999" }),
                "找不到");
            ExpectFail(run, "wt-opt-unk",
                s.OptimizeTableDisplay(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = "T_99999" }),
                "找不到");

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            Ok(run, "wt-two-del-first",
                del.Success && !StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                del.Success ? "删第一张时第二张也没了或第一张还在" : (del.Error ?? ""));
            Ok(run, "wt-del-keep-sent",
                del.Success && StillHas(s, id, ConsSent),
                del.Success ? "删表带走了巩固句" : (del.Error ?? ""));
            ExpectFail(run, "wt-read-gone",
                s.ReadTableRowValues(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = tableId
                }),
                "找不到");
            ExpectFail(run, "wt-del-twice",
                s.DeleteTable(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = tableId }),
                "找不到");
            ExpectFail(run, "wt-opt-after-del",
                s.OptimizeTableDisplay(new Dictionary<string, object> { ["channel_id"] = id, ["table_id"] = tableId }),
                "找不到");
            ExpectFail(run, "wt-fill-del",
                s.ApplyTableRowValues(FillArgs(id, tableId, 1, new[] { "x" })),
                "找不到");

            ToolResult rebuild = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(RebuildA, "1", RebuildB, "2"))));
            string newId = WdSession.Str(rebuild, "table_id");
            if (!rebuild.Success || string.IsNullOrEmpty(newId))
            {
                FailMany(run, rebuild.Error ?? "巩固重建失败", "wt-opt-rebuild");
            }
            else
            {
                ToolResult opt = s.OptimizeTableDisplay(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["table_id"] = newId
                });
                Ok(run, "wt-opt-rebuild", opt.Success, opt.Error ?? "新表第一次 optimize 失败");
            }

            s.CloseDocByPath(path);
        }

        private static void RunConsolidateMore(TestRun run, WdSession s)
        {
            RunMoreSolo(run, s);
            RunMoreTwoSent(run, s);
            RunMoreBeforeAndCursor(run, s);
            RunMoreDelFirstKeepSecond(run, s);
            RunMoreShapesAndNeg(run, s);
        }

        private static void RunMoreSolo(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ExtraIds2))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                FailMany(run, "COM 写单句失败",
                    "wt-solo-after", "wt-solo-after-id", "wt-solo-fill", "wt-solo-fill-text", "wt-solo-del-keep");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string sent = CodeContaining(ReadBody(s, id), AfterSent, "S");
            var afterArgs = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            afterArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult created = s.CreateTableFromXml(afterArgs);
            string tableId = WdSession.Str(created, "table_id");
            Ok(run, "wt-solo-after",
                created.Success && StillHas(s, id, AfterSent) && StillHas(s, id, ConsA),
                created.Success ? "单句后建表丢了句或表" : (created.Error ?? ""));
            Ok(run, "wt-solo-after-id",
                created.Success && tableId.StartsWith("T_", StringComparison.Ordinal),
                created.Success ? "单句后无 T_" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(tableId))
            {
                FailMany(run, created.Error ?? "单句后无 table_id",
                    "wt-solo-fill", "wt-solo-fill-text", "wt-solo-del-keep");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult read = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            int emptyRow, emptyCount;
            if (!read.Success || !TryFirstEmptyRow(read, out emptyRow, out emptyCount))
            {
                FailMany(run, read.Success ? "单句表无空槽" : (read.Error ?? ""),
                    "wt-solo-fill", "wt-solo-fill-text", "wt-solo-del-keep");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult fill = s.ApplyTableRowValues(
                FillArgs(id, tableId, emptyRow, BuildApplyValues(read, emptyRow, FillA)));
            Ok(run, "wt-solo-fill", fill.Success, fill.Error ?? "单句表填空失败");
            Ok(run, "wt-solo-fill-text",
                fill.Success && StillHas(s, id, FillA) && StillHas(s, id, ConsA) && StillHas(s, id, AfterSent),
                fill.Success ? "单句表填空字不对" : (fill.Error ?? ""));

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = tableId
            });
            Ok(run, "wt-solo-del-keep",
                del.Success && !StillHas(s, id, ConsA) && StillHas(s, id, AfterSent),
                del.Success ? "删单句表带走了前句或标签还在" : (del.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunMoreTwoSent(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id,
                "wt-two-after-first", "wt-two-keep-both", "wt-two-order",
                "wt-two-after-last", "wt-two-last-order"))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent + "\r" + AfterSentB))
            {
                FailMany(run, "COM 写两句失败",
                    "wt-two-after-first", "wt-two-keep-both", "wt-two-order",
                    "wt-two-after-last", "wt-two-last-order");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string body = ReadBody(s, id);
            string first = CodeContaining(body, AfterSent, "S");
            var firstArgs = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            firstArgs["target"] = new Dictionary<string, object> { ["codes"] = first };
            ToolResult afterFirst = s.CreateTableFromXml(firstArgs);
            string afterBody = ReadBody(s, id);
            Ok(run, "wt-two-after-first",
                afterFirst.Success && StillHas(s, id, ConsA),
                afterFirst.Success ? "第一句后无表" : (afterFirst.Error ?? ""));
            Ok(run, "wt-two-keep-both",
                afterFirst.Success && StillHas(s, id, AfterSent) && StillHas(s, id, AfterSentB),
                afterFirst.Success ? "第一句后建表丢了句子" : (afterFirst.Error ?? ""));
            Ok(run, "wt-two-order",
                afterFirst.Success && InOrder(afterBody, AfterSent, ConsA, AfterSentB),
                afterFirst.Success ? "第一句后表不在两句之间" : (afterFirst.Error ?? ""));
            s.CloseDocByPath(path);

            if (!TryOpen(run, s, out path, out id, "wt-two-after-last", "wt-two-last-order"))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent + "\r" + AfterSentB))
            {
                FailMany(run, "COM 写两句(末)失败", "wt-two-after-last", "wt-two-last-order");
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            string last = CodeContaining(ReadBody(s, id), AfterSentB, "S");
            var lastArgs = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            lastArgs["target"] = new Dictionary<string, object> { ["codes"] = last };
            ToolResult afterLast = s.CreateTableFromXml(lastArgs);
            string lastBody = ReadBody(s, id);
            Ok(run, "wt-two-after-last",
                afterLast.Success && StillHas(s, id, ConsC) && StillHas(s, id, AfterSentB),
                afterLast.Success ? "末句后建表失败" : (afterLast.Error ?? ""));
            Ok(run, "wt-two-last-order",
                afterLast.Success && InOrder(lastBody, AfterSent, AfterSentB, ConsC),
                afterLast.Success ? "末句后表不在两句之后" : (afterLast.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunMoreBeforeAndCursor(TestRun run, WdSession s)
        {
            string[] rest =
            {
                "wt-before-tbl", "wt-before-tbl-id", "wt-before-keep-sent", "wt-before-order",
                "wt-read-first-after-before", "wt-third-after-tbl", "wt-third-text",
                "wt-cursor-after-two", "wt-cursor-first-text", "wt-cursor-n", "wt-no-join-labels",
                "wt-fill-first-late", "wt-fill-first-late-text", "wt-read-second", "wt-fill-second",
                "wt-fill-second-text", "wt-opt-second", "wt-opt-first-late", "wt-xfmt-second",
                "wt-afmt-first-late", "wt-afmt-second-keep", "wt-del-second-keep-first",
                "wt-start-keep-old", "wt-fill-second-short", "wt-create-after-filled"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, rest))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                FailMany(run, "COM 写表前句失败", rest);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.MoveSelectionToEnd(path);
            ToolResult first = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, ""))));
            string firstId = WdSession.Str(first, "table_id");
            if (!first.Success || string.IsNullOrEmpty(firstId))
            {
                FailMany(run, first.Error ?? "表前铺甲失败", rest);
                s.CloseDocByPath(path);
                return;
            }

            string sent = CodeContaining(ReadBody(s, id), AfterSent, "S");
            var beforeArgs = CreateArgs(id, WriteXml(GridXml(ConsC, "", ConsD, "")));
            beforeArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
            ToolResult second = s.CreateTableFromXml(beforeArgs);
            string secondId = WdSession.Str(second, "table_id");
            string beforeBody = ReadBody(s, id);
            Ok(run, "wt-before-tbl",
                second.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                second.Success ? "表前建表合成或丢表" : (second.Error ?? ""));
            Ok(run, "wt-before-tbl-id",
                second.Success && secondId.StartsWith("T_", StringComparison.Ordinal) && secondId != firstId,
                second.Success ? "表前第二张 id 不对" : (second.Error ?? ""));
            Ok(run, "wt-before-keep-sent",
                second.Success && StillHas(s, id, AfterSent),
                second.Success ? "表前建表丢了前句" : (second.Error ?? ""));
            Ok(run, "wt-before-order",
                second.Success && InOrder(beforeBody, AfterSent, ConsC, ConsA),
                second.Success ? "表前表不在句与甲之间" : (second.Error ?? ""));
            if (!second.Success || string.IsNullOrEmpty(secondId))
            {
                FailMany(run, second.Error ?? "表前无 table_id",
                    "wt-read-first-after-before", "wt-third-after-tbl", "wt-third-text",
                    "wt-cursor-after-two", "wt-cursor-first-text", "wt-cursor-n", "wt-no-join-labels",
                    "wt-fill-first-late", "wt-fill-first-late-text", "wt-read-second", "wt-fill-second",
                    "wt-fill-second-text", "wt-opt-second", "wt-opt-first-late", "wt-xfmt-second",
                    "wt-afmt-first-late", "wt-afmt-second-keep", "wt-del-second-keep-first",
                    "wt-start-keep-old", "wt-fill-second-short", "wt-create-after-filled");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult readFirst = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = firstId
            });
            int firstEmptyRow = 0;
            int firstEmpty = 0;
            bool firstHasEmpty = readFirst.Success && TryFirstEmptyRow(readFirst, out firstEmptyRow, out firstEmpty);
            Ok(run, "wt-read-first-after-before",
                firstHasEmpty,
                readFirst.Success ? "表前插入后甲空槽没了" : (readFirst.Error ?? ""));

            var thirdArgs = CreateArgs(id, WriteXml(GridXml(ConsE, "", Cell1, "")));
            thirdArgs["target"] = new Dictionary<string, object> { ["table"] = secondId };
            ToolResult third = s.CreateTableFromXml(thirdArgs);
            string thirdId = WdSession.Str(third, "table_id");
            Ok(run, "wt-third-after-tbl",
                third.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC) && StillHas(s, id, ConsE),
                third.Success ? "第三张表后建表连表" : (third.Error ?? ""));
            Ok(run, "wt-third-text",
                third.Success && thirdId != firstId && thirdId != secondId,
                third.Success ? "第三张 table_id 不对" : (third.Error ?? ""));

            ToolResult cursor1 = s.CreateTableFromXml(CreateArgs(id, WriteXml(OneCellXml(Row3A))));
            Ok(run, "wt-cursor-after-two",
                cursor1.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC) && StillHas(s, id, Row3A),
                cursor1.Success ? "光标再建吃掉前表" : (cursor1.Error ?? ""));
            ToolResult readAfterCursor = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = firstId
            });
            int afterCursorEmpty;
            Ok(run, "wt-cursor-first-text",
                readAfterCursor.Success && TryFirstEmptyRow(readAfterCursor, out _, out afterCursorEmpty),
                readAfterCursor.Success ? "光标再建后甲空槽没了" : (readAfterCursor.Error ?? ""));

            ToolResult cursorN = s.CreateTableFromXml(CreateArgs(id, WriteXml(ThreeRowXml())));
            Ok(run, "wt-cursor-n",
                cursorN.Success && StillHas(s, id, ConsA) && StillHas(s, id, Row3B) && StillHas(s, id, Row3C),
                cursorN.Success ? "光标三行表连进甲" : (cursorN.Error ?? ""));
            ToolResult cursorRs = s.CreateTableFromXml(CreateArgs(id, WriteXml(RowspanXml())));
            string afterJoin = ReadBody(s, id);
            Ok(run, "wt-no-join-labels",
                cursorRs.Success
                    && InOrder(afterJoin, ConsC, ConsA)
                    && afterJoin.IndexOf(ConsA, StringComparison.Ordinal) >= 0
                    && afterJoin.IndexOf(ConsC, StringComparison.Ordinal) >= 0
                    && afterJoin.IndexOf(RsLeft, StringComparison.Ordinal) >= 0
                    && afterJoin.IndexOf(Row3A, StringComparison.Ordinal) >= 0,
                cursorRs.Success ? "连建后标签被覆盖" : (cursorRs.Error ?? ""));

            if (!firstHasEmpty)
            {
                FailMany(run, "甲无空槽，无法晚填",
                    "wt-fill-first-late", "wt-fill-first-late-text", "wt-create-after-filled");
            }
            else
            {
                ToolResult fillFirst = s.ApplyTableRowValues(
                    FillArgs(id, firstId, firstEmptyRow, BuildApplyValues(readFirst, firstEmptyRow, FillA)));
                Ok(run, "wt-fill-first-late", fillFirst.Success, fillFirst.Error ?? "晚填甲失败");
                Ok(run, "wt-fill-first-late-text",
                    fillFirst.Success && StillHas(s, id, FillA) && StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                    fillFirst.Success ? "晚填甲字不对" : (fillFirst.Error ?? ""));

                var afterFilled = CreateArgs(id, WriteXml(GridXml(MoreStart, "1", PairA, "2")));
                afterFilled["target"] = new Dictionary<string, object> { ["table"] = firstId };
                ToolResult filledAfter = s.CreateTableFromXml(afterFilled);
                Ok(run, "wt-create-after-filled",
                    filledAfter.Success && StillHas(s, id, FillA) && StillHas(s, id, MoreStart) && StillHas(s, id, ConsA),
                    filledAfter.Success ? "已填表后再建表丢了字" : (filledAfter.Error ?? ""));
            }

            ToolResult readSecond = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = secondId
            });
            int secondRow = 0;
            int secondEmpty = 0;
            bool secondHasEmpty = readSecond.Success && TryFirstEmptyRow(readSecond, out secondRow, out secondEmpty);
            Ok(run, "wt-read-second",
                secondHasEmpty,
                readSecond.Success ? "丙无空槽" : (readSecond.Error ?? ""));
            if (!secondHasEmpty)
            {
                FailMany(run, readSecond.Success ? "丙无空槽" : (readSecond.Error ?? ""),
                    "wt-fill-second", "wt-fill-second-text", "wt-fill-second-short");
            }
            else
            {
                int secondSlots;
                TryRowSlotCount(readSecond, secondRow, out secondSlots);
                var tooShort = RepeatSlots(Math.Max(0, secondSlots - 1), "");

                ToolResult shortLen = s.ApplyTableRowValues(FillArgs(id, secondId, secondRow, tooShort));
                Ok(run, "wt-fill-second-short",
                    !shortLen.Success && StillHas(s, id, ConsC) && !StillHas(s, id, FillB),
                    shortLen.Success ? "丙 values 偏短仍写入" : (shortLen.Error ?? "错长度却改了丙"));

                ToolResult fillSecond = s.ApplyTableRowValues(
                    FillArgs(id, secondId, secondRow, BuildApplyValues(readSecond, secondRow, FillB)));
                Ok(run, "wt-fill-second", fillSecond.Success, fillSecond.Error ?? "填丙失败");
                Ok(run, "wt-fill-second-text",
                    fillSecond.Success && StillHas(s, id, FillB) && StillHas(s, id, ConsC) && StillHas(s, id, ConsA),
                    fillSecond.Success ? "填丙字不对" : (fillSecond.Error ?? ""));
            }

            ToolResult optSecond = s.OptimizeTableDisplay(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = secondId
            });
            Ok(run, "wt-opt-second", optSecond.Success, optSecond.Error ?? "优丙失败");
            ToolResult optFirst = s.OptimizeTableDisplay(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = firstId
            });
            Ok(run, "wt-opt-first-late", optFirst.Success, optFirst.Error ?? "晚优甲失败");

            string extractPath = Path.Combine(Path.GetTempPath(), "ew-wt-x2-" + Guid.NewGuid().ToString("N") + ".xml");
            ToolResult xfmt = s.ExtractTableFormat(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["path"] = extractPath,
                ["table_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["table_id"] = secondId
                }
            });
            Ok(run, "wt-xfmt-second",
                xfmt.Success && File.Exists(extractPath),
                xfmt.Success ? "抽丙格式无文件" : (xfmt.Error ?? ""));

            string fmtXml =
                "<TableConfig><General><TableFontColor>000000</TableFontColor><Style><TableStyle>ThreeLine</TableStyle><HeaderRows>1</HeaderRows></Style></General></TableConfig>";
            ToolResult afmtFirst = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = firstId,
                ["xml_content"] = fmtXml
            });
            Ok(run, "wt-afmt-first-late",
                afmtFirst.Success && StillHas(s, id, ConsA),
                afmtFirst.Success ? "晚套甲改了字" : (afmtFirst.Error ?? ""));
            ToolResult afmtSecond = s.ApplyTableFormatFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = secondId,
                ["xml_content"] = fmtXml
            });
            Ok(run, "wt-afmt-second-keep",
                afmtSecond.Success && StillHas(s, id, ConsA) && StillHas(s, id, ConsC),
                afmtSecond.Success ? "套丙带走了甲" : (afmtSecond.Error ?? ""));

            var startArgs = CreateArgs(id, WriteXml(OneCellXml(MoreStart + "首")));
            startArgs["target"] = new Dictionary<string, object> { ["position"] = "start" };
            ToolResult atStart = s.CreateTableFromXml(startArgs);
            Ok(run, "wt-start-keep-old",
                atStart.Success && StillHas(s, id, ConsA) && StillHas(s, id, MoreStart + "首"),
                atStart.Success ? "文首建表冲掉旧表" : (atStart.Error ?? ""));

            ToolResult delSecond = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = secondId
            });
            Ok(run, "wt-del-second-keep-first",
                delSecond.Success && !StillHas(s, id, ConsC) && StillHas(s, id, ConsA),
                delSecond.Success ? "删丙时甲也没了或丙还在" : (delSecond.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunMoreDelFirstKeepSecond(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, "wt-del-first-keep-second"))
            {
                return;
            }

            ToolResult a = s.CreateTableFromXml(CreateArgs(id, WriteXml(GridXml(PairA, "1", PairB, "2"))));
            string aId = WdSession.Str(a, "table_id");
            if (!a.Success || string.IsNullOrEmpty(aId))
            {
                FailMany(run, a.Error ?? "对表甲失败", "wt-del-first-keep-second");
                s.CloseDocByPath(path);
                return;
            }

            var bArgs = CreateArgs(id, WriteXml(OneCellXml(TallA)));
            bArgs["target"] = new Dictionary<string, object> { ["table"] = aId };
            ToolResult b = s.CreateTableFromXml(bArgs);
            string bId = WdSession.Str(b, "table_id");
            if (!b.Success || string.IsNullOrEmpty(bId))
            {
                FailMany(run, b.Error ?? "对表乙失败", "wt-del-first-keep-second");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult del = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = aId
            });
            Ok(run, "wt-del-first-keep-second",
                del.Success && !StillHas(s, id, PairA) && StillHas(s, id, TallA),
                del.Success ? "删第一张时第二张也没了" : (del.Error ?? ""));
            s.CloseDocByPath(path);
        }

        private static void RunMoreShapesAndNeg(TestRun run, WdSession s)
        {
            string[] rest =
            {
                "wt-create-unk-chart", "wt-create-unk-image", "wt-create-two-modes",
                "wt-create-empty-codes", "wt-create-empty-table",
                "wt-create-2x3", "wt-create-3x2", "wt-create-2x1", "wt-read-2x3-rows",
                "wt-fill-2x3", "wt-opt-2x3", "wt-del-2x3-gone",
                "wt-create-bad-merge", "wt-create-empty-tr"
            };
            string path, id;
            if (!TryOpen(run, s, out path, out id, rest))
            {
                return;
            }

            var unkChart = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkChart["target"] = new Dictionary<string, object> { ["chart"] = "C_99999" };
            ExpectFail(run, "wt-create-unk-chart", s.CreateTableFromXml(unkChart), "图表");

            var unkImage = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            unkImage["target"] = new Dictionary<string, object> { ["image"] = "I_99999" };
            ExpectFail(run, "wt-create-unk-image", s.CreateTableFromXml(unkImage), "图片");

            var twoModes = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            twoModes["target"] = new Dictionary<string, object>
            {
                ["codes"] = "S_00001",
                ["table"] = "T_00001"
            };
            ExpectFail(run, "wt-create-two-modes", s.CreateTableFromXml(twoModes), "六选一");

            var emptyCodes = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            emptyCodes["target"] = new Dictionary<string, object> { ["codes"] = "   " };
            ExpectFail(run, "wt-create-empty-codes", s.CreateTableFromXml(emptyCodes), "六选一");

            var emptyTable = CreateArgs(id, WriteXml(GridXml(ConsA, "", ConsB, "")));
            emptyTable["target"] = new Dictionary<string, object> { ["table"] = "" };
            ExpectFail(run, "wt-create-empty-table", s.CreateTableFromXml(emptyTable), "六选一");

            ExpectFail(run, "wt-create-bad-merge",
                s.CreateTableFromXml(CreateArgs(id, WriteXml(
                    "<TableConfig><General/><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                    + "<tr><td>甲</td><td>乙</td></tr>"
                    + "<tr><td colspan=\"3\">越界</td></tr></Data></Table></TableConfig>"))),
                "越界");
            ToolResult emptyTr = s.CreateTableFromXml(CreateArgs(id, WriteXml(
                "<TableConfig><General/><Table><Properties/><Data><tr></tr></Data></Table></TableConfig>")));
            Ok(run, "wt-create-empty-tr",
                emptyTr != null && !emptyTr.Success
                    && (Contains(emptyTr.Error, "为空") || Contains(emptyTr.Error, "推断列数")),
                emptyTr != null && !emptyTr.Success
                    ? "失败文案不含「为空/推断列数」: " + (emptyTr.Error ?? "")
                    : "空 <tr> 仍成功");

            ToolResult wide = s.CreateTableFromXml(CreateArgs(id, WriteXml(WideXml())));
            string wideId = WdSession.Str(wide, "table_id");
            Ok(run, "wt-create-2x3",
                wide.Success && StillHas(s, id, WideA) && StillHas(s, id, WideB) && StillHas(s, id, WideC),
                wide.Success ? "2x3 缺标签" : (wide.Error ?? ""));

            ToolResult tall = s.CreateTableFromXml(CreateArgs(id, WriteXml(TallXml())));
            Ok(run, "wt-create-3x2",
                tall.Success && StillHas(s, id, TallA) && StillHas(s, id, TallB),
                tall.Success ? "3x2 缺标签" : (tall.Error ?? ""));

            ToolResult pair = s.CreateTableFromXml(CreateArgs(id, WriteXml(PairColXml())));
            Ok(run, "wt-create-2x1",
                pair.Success && StillHas(s, id, PairA) && StillHas(s, id, PairB),
                pair.Success ? "2x1 缺标签" : (pair.Error ?? ""));

            if (!wide.Success || string.IsNullOrEmpty(wideId))
            {
                FailMany(run, wide.Error ?? "2x3 无 id",
                    "wt-read-2x3-rows", "wt-fill-2x3", "wt-opt-2x3", "wt-del-2x3-gone");
                s.CloseDocByPath(path);
                return;
            }

            ToolResult readWide = s.ReadTableRowValues(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = wideId
            });
            JObject data = DataObj(readWide);
            JArray rows = data == null ? null : data["rows"] as JArray;
            Ok(run, "wt-read-2x3-rows",
                readWide.Success && rows != null && rows.Count == 2,
                readWide.Success ? "2x3 rows=" + (rows == null ? -1 : rows.Count) : (readWide.Error ?? ""));

            int wideRow, wideEmpty;
            if (!readWide.Success || !TryFirstEmptyRow(readWide, out wideRow, out wideEmpty))
            {
                FailMany(run, readWide.Success ? "2x3 无空槽" : (readWide.Error ?? ""), "wt-fill-2x3");
            }
            else
            {
                ToolResult fillWide = s.ApplyTableRowValues(
                    FillArgs(id, wideId, wideRow, BuildApplyValues(readWide, wideRow, FillA)));
                Ok(run, "wt-fill-2x3",
                    fillWide.Success && StillHas(s, id, FillA) && StillHas(s, id, WideA),
                    fillWide.Success ? "2x3 填空字不对" : (fillWide.Error ?? ""));
            }

            ToolResult optWide = s.OptimizeTableDisplay(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = wideId
            });
            Ok(run, "wt-opt-2x3", optWide.Success, optWide.Error ?? "2x3 optimize 失败");

            ToolResult delWide = s.DeleteTable(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = wideId
            });
            Ok(run, "wt-del-2x3-gone",
                delWide.Success && !StillHas(s, id, WideA) && StillHas(s, id, TallA) && StillHas(s, id, PairA),
                delWide.Success ? "删 2x3 时其他表也没了或标签还在" : (delWide.Error ?? ""));
            s.CloseDocByPath(path);
        }

        public static void RunNestedAfterOnly(
            TestRun run,
            string host,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            var s = new WdSession(host);
            RunNestedTableAfter(run, s);
            RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
        }

        private static void RunNestedTableAfter(TestRun run, WdSession s)
        {
            const string OuterCell = "易写外表格。";
            const string InnerA = "易写内表甲";
            const string InnerB = "易写内表乙";
            const string NestedAfter = "易写嵌套表后句。";
            var ids = new[] { "wt-nested-create-id", "wt-nested-after-ins", "wt-nested-after-in-cell" };
            string path, id;
            if (!TryOpen(run, s, out path, out id, ids))
            {
                return;
            }

            try
            {
                s.ActivateDoc(path);
                ToolResult outer = s.CreateTableFromXml(CreateArgs(id, WriteXml(OneCellXml(OuterCell))));
                string outerId = WdSession.Str(outer, "table_id");
                if (!outer.Success || string.IsNullOrEmpty(outerId))
                {
                    FailMany(run, outer.Error ?? "外表建表失败", ids);
                    return;
                }

                string sent = CodeContaining(ReadBody(s, id), OuterCell, "S");
                if (string.IsNullOrEmpty(sent))
                {
                    FailMany(run, "外表格无 S_", ids);
                    return;
                }

                var innerArgs = CreateArgs(id, WriteXml(GridXml(InnerA, InnerB, "备", "")));
                innerArgs["target"] = new Dictionary<string, object> { ["codes"] = sent };
                ToolResult inner = s.CreateTableFromXml(innerArgs);
                string returnedId = WdSession.Str(inner, "table_id");
                string body = ReadBody(s, id);
                List<string> tableIds = CollectTableIds(body);
                string displayInner = null;
                for (int i = 0; i < tableIds.Count; i++)
                {
                    if (!string.Equals(tableIds[i], outerId, StringComparison.Ordinal))
                    {
                        displayInner = tableIds[i];
                        break;
                    }
                }
                Ok(run, "wt-nested-create-id",
                    inner.Success
                        && !string.IsNullOrEmpty(returnedId)
                        && returnedId != outerId
                        && returnedId == displayInner,
                    inner.Success
                        ? ("返回 " + returnedId + " 外表 " + outerId + " display=" + string.Join(",", tableIds))
                        : (inner.Error ?? ""));
                if (!inner.Success || string.IsNullOrEmpty(returnedId) || returnedId != displayInner)
                {
                    FailMany(run, inner.Error ?? "内表 table_id 与读文不一致",
                        "wt-nested-after-ins", "wt-nested-after-in-cell");
                    return;
                }

                ToolResult ins = s.ProcessActions(id, new List<Dictionary<string, object>>
                {
                    new Dictionary<string, object>
                    {
                        ["type"] = "insert",
                        ["detail"] = new Dictionary<string, object>
                        {
                            ["target"] = new Dictionary<string, object> { ["table"] = returnedId },
                            ["insert"] = NestedAfter,
                            ["format"] = DefaultFormat()
                        }
                    }
                });
                string afterBody = ReadBody(s, id);
                Ok(run, "wt-nested-after-ins",
                    ins.Success && afterBody.IndexOf(NestedAfter, StringComparison.Ordinal) >= 0,
                    ins.Success ? "表后句未出现" : (ins.Error ?? ""));

                int innerTag = afterBody.IndexOf("id=\"" + returnedId + "\"", StringComparison.Ordinal);
                int innerClose = innerTag < 0 ? -1 : afterBody.IndexOf("</table>", innerTag, StringComparison.Ordinal);
                int textAt = innerClose < 0 ? -1 : afterBody.IndexOf(NestedAfter, innerClose, StringComparison.Ordinal);
                int outerClose = afterBody.LastIndexOf("</table>", StringComparison.Ordinal);
                bool inParent = textAt > innerClose && textAt < outerClose;
                Ok(run, "wt-nested-after-in-cell",
                    ins.Success && inParent,
                    ins.Success ? "表后句不在父格内（内表后、外表结束前）" : (ins.Error ?? ""));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static List<string> CollectTableIds(string body)
        {
            var ids = new List<string>();
            if (string.IsNullOrEmpty(body))
            {
                return ids;
            }

            MatchCollection matches = Regex.Matches(body, "<table id=\"(T_[^\"]+)\"");
            for (int i = 0; i < matches.Count; i++)
            {
                ids.Add(matches[i].Groups[1].Value);
            }

            return ids;
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string xml = WriteXml(GridXml(LabelA, "1", LabelB, "2"));
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wt-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wt-neg-wrong-kind-excel",
                    s.CreateTableFromXml(CreateArgs(WdSession.Str(xl, "channel_id"), xml)), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wt-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wt-neg-wrong-kind-ppt",
                    s.CreateTableFromXml(CreateArgs(WdSession.Str(ppt, "channel_id"), xml)), "ppt");
                WdSession.ClosePptByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wt-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wt-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.CreateTableFromXml(CreateArgs(cid, xml));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wt-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wt-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wt-neg-closed", "关稿后仍成功");
            }
        }

        private static void RunNoQuit(
            TestRun run,
            int wordBefore,
            int wpsBefore,
            bool wordWinBefore,
            bool wpsWinBefore)
        {
            string wordDetail;
            string wpsDetail;
            bool wordOk = OfficeAlive.NotQuit("WINWORD", wordBefore, wordWinBefore, out wordDetail);
            bool wpsOk = OfficeAlive.NotQuit("wps", wpsBefore, wpsWinBefore, out wpsDetail);
            Ok(run, "wt-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
        }

        private static bool TryOpen(
            TestRun run,
            WdSession s,
            out string path,
            out string id,
            params string[] failIds)
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
            return true;
        }

        private static Dictionary<string, object> CreateArgs(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["path"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static Dictionary<string, object> CreateArgsFilename(string channelId, string path)
        {
            var args = new Dictionary<string, object> { ["filename"] = path ?? "" };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
        }

        private static bool CreateOkHas(WdSession s, string id, string xml, string needle)
        {
            ToolResult created = s.CreateTableFromXml(CreateArgs(id, WriteXml(xml)));
            return created.Success && StillHas(s, id, needle);
        }

        private static string[] RepeatSlots(int count, string text)
        {
            var values = new string[Math.Max(0, count)];
            for (int i = 0; i < values.Length; i++)
            {
                values[i] = text ?? "";
            }

            return values;
        }

        private static bool TryRowSlotCount(ToolResult read, int row, out int slotCount)
        {
            slotCount = 0;
            JObject rowObj = FindRowObject(read, row);
            if (rowObj == null)
            {
                return false;
            }

            slotCount = RowSlotCount(rowObj);
            return true;
        }

        private static JObject FindRowObject(ToolResult read, int row)
        {
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return null;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                try
                {
                    if (o["row"] != null && o["row"].Value<int>() == row)
                    {
                        return o;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static int RowSlotCount(JObject rowObj)
        {
            if (rowObj == null)
            {
                return 0;
            }

            JArray slots = rowObj["slots"] as JArray;
            if (slots != null)
            {
                return slots.Count;
            }

            JArray values = rowObj["values"] as JArray;
            return values == null ? 0 : values.Count;
        }

        private static string[] BuildApplyValues(ToolResult read, int row, string firstEmptyFill)
        {
            JObject rowObj = FindRowObject(read, row);
            if (rowObj == null)
            {
                return new string[0];
            }

            JArray slots = rowObj["slots"] as JArray;
            if (slots == null)
            {
                return RepeatSlots(RowSlotCount(rowObj), "<keep/>");
            }

            var values = new string[slots.Count];
            bool filledFirst = false;
            for (int i = 0; i < slots.Count; i++)
            {
                JObject slot = slots[i] as JObject;
                bool empty = false;
                try
                {
                    empty = slot != null && slot["empty"] != null && slot["empty"].Value<bool>();
                }
                catch
                {
                    empty = false;
                }

                if (!empty)
                {
                    values[i] = "<keep/>";
                    continue;
                }

                if (!filledFirst && !string.IsNullOrEmpty(firstEmptyFill))
                {
                    values[i] = firstEmptyFill;
                    filledFirst = true;
                }
                else
                {
                    values[i] = "";
                }
            }

            return values;
        }

        private static string OneCellXml(string text)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + text + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string ThreeRowXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + Row3A + "</td></tr>"
                + "<tr><td>" + Row3B + "</td></tr>"
                + "<tr><td>" + Row3C + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string SpanXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td><span style='font-weight:bold'>" + SpanTxt + "</span></td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string RowspanXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td rowspan=\"2\">" + RsLeft + "</td><td>" + RsUp + "</td></tr>"
                + "<tr><td>" + RsDown + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static bool RowHasSlots(JArray rows)
        {
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o != null && o["slots"] is JArray)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool TryLastEmptyRow(ToolResult read, out int row, out int emptyCount)
        {
            row = 0;
            emptyCount = 0;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            bool found = false;
            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int n = 0;
                try
                {
                    n = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                }
                catch
                {
                    n = 0;
                }

                if (n > 0)
                {
                    try
                    {
                        row = o["row"].Value<int>();
                    }
                    catch
                    {
                        return false;
                    }

                    emptyCount = n;
                    found = true;
                }
            }

            return found;
        }

        private static bool TryRowEmptyCount(ToolResult read, int row, out int emptyCount)
        {
            emptyCount = -1;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int r;
                try
                {
                    r = o["row"].Value<int>();
                }
                catch
                {
                    continue;
                }

                if (r != row)
                {
                    continue;
                }

                try
                {
                    emptyCount = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        private static Dictionary<string, object> FillArgs(
            string channelId,
            string tableId,
            int row,
            IList<string> values)
        {
            var args = new Dictionary<string, object>
            {
                ["table_id"] = tableId ?? "",
                ["format"] = DefaultFormat(),
                ["writes"] = new object[]
                {
                    new Dictionary<string, object>
                    {
                        ["row"] = row,
                        ["values"] = values
                    }
                }
            };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            return args;
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

        private static bool InOrder(string body, params string[] parts)
        {
            if (string.IsNullOrEmpty(body) || parts == null)
            {
                return false;
            }

            int pos = 0;
            foreach (string part in parts)
            {
                if (string.IsNullOrEmpty(part))
                {
                    return false;
                }

                int i = body.IndexOf(part, pos, StringComparison.Ordinal);
                if (i < 0)
                {
                    return false;
                }

                pos = i + part.Length;
            }

            return true;
        }

        private static string WideXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties>"
                + "<ColWidths><float>80</float><float>80</float><float>80</float></ColWidths></Properties><Data>"
                + "<tr><td>" + WideA + "</td><td></td><td>" + WideB + "</td></tr>"
                + "<tr><td>" + WideC + "</td><td></td><td></td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string TallXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties>"
                + "<ColWidths><float>120</float><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + TallA + "</td><td></td></tr>"
                + "<tr><td>" + TallB + "</td><td></td></tr>"
                + "<tr><td></td><td></td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string PairColXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties>"
                + "<ColWidths><float>120</float></ColWidths></Properties><Data>"
                + "<tr><td>" + PairA + "</td></tr>"
                + "<tr><td>" + PairB + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string GridXml(string a, string b, string c, string d)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td>" + a + "</td><td>" + b + "</td></tr>"
                + "<tr><td>" + c + "</td><td>" + d + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string MergeXml()
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths><float>120</float><float>180</float></ColWidths></Properties><Data>"
                + "<tr><td colspan=\"2\">" + MergeHead + "</td></tr>"
                + "<tr><td>" + MergeLeft + "</td><td>" + MergeRight + "</td></tr>"
                + "</Data></Table></TableConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wt-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
        }

        private static bool TryFirstEmptyRow(ToolResult read, out int row, out int emptyCount)
        {
            row = 0;
            emptyCount = 0;
            JObject data = DataObj(read);
            JArray rows = data == null ? null : data["rows"] as JArray;
            if (rows == null)
            {
                return false;
            }

            foreach (JToken token in rows)
            {
                JObject o = token as JObject;
                if (o == null)
                {
                    continue;
                }

                int n = 0;
                try
                {
                    n = o["empty_count"] == null ? 0 : o["empty_count"].Value<int>();
                }
                catch
                {
                    n = 0;
                }

                if (n > 0)
                {
                    try
                    {
                        row = o["row"].Value<int>();
                    }
                    catch
                    {
                        return false;
                    }

                    emptyCount = n;
                    return true;
                }
            }

            return false;
        }

        private static string FirstCode(string text, string prefix)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            Match m = Regex.Match(text, @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\]");
            return m.Success ? prefix + "_" + m.Groups[1].Value : "";
        }

        private static string CodeContaining(string text, string needle, string prefix)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle))
            {
                return "";
            }

            MatchCollection matches = Regex.Matches(
                text,
                @"\[" + prefix + @"_([0-9a-zA-Z]{5})-start\](.*?)\[\s*" + prefix + @"_\1-end\]",
                RegexOptions.Singleline);
            foreach (Match m in matches)
            {
                if (m.Groups[2].Value.IndexOf(needle, StringComparison.Ordinal) >= 0)
                {
                    return prefix + "_" + m.Groups[1].Value;
                }
            }

            return "";
        }

        private static string ReadBody(WdSession s, string id)
        {
            ToolResult got = s.GetContent(id);
            return got.Success ? WdSession.Str(got, "document_content") : "";
        }

        private static bool StillHas(WdSession s, string id, string needle)
        {
            return ReadBody(s, id).IndexOf(needle, StringComparison.Ordinal) >= 0;
        }

        private static JObject DataObj(ToolResult r)
        {
            if (r == null || r.Data == null)
            {
                return null;
            }

            try
            {
                return JObject.Parse(JsonConvert.SerializeObject(r.Data));
            }
            catch
            {
                return null;
            }
        }

        private static void ExpectFail(TestRun run, string id, ToolResult r, string needle)
        {
            if (r != null && !r.Success && (Contains(r.Error, needle) || BlobHas(r, needle)))
            {
                run.CaseOk(id);
            }
            else if (r != null && !r.Success)
            {
                run.CaseFail(id, "失败文案不含「" + needle + "」: " + (r.Error ?? ""));
            }
            else
            {
                run.CaseFail(id, "应失败");
            }
        }

        private static bool BlobHas(ToolResult r, string needle)
        {
            if (Contains(r == null ? null : r.Error, needle))
            {
                return true;
            }

            if (r == null || r.Data == null)
            {
                return false;
            }

            try
            {
                return Contains(JsonConvert.SerializeObject(r.Data), needle);
            }
            catch
            {
                return Contains(Convert.ToString(r.Data), needle);
            }
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
    }
}
