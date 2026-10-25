using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>多格几何码路径巩固 50 条：合并格、表外混批、两表、删/换/插。</summary>
    internal static class WmConsolidate
    {
        public const int Count = 50;

        private const string A = "固甲。";
        private const string B = "固乙。";
        private const string C = "固丙。";
        private const string Outside = "表前句。";
        private const string After = "表后句。";

        public static void Run(TestRun run, WdSession s)
        {
            RunColspan(run, s);
            RunRowspan(run, s);
            RunRowCol(run, s);
            RunOutside(run, s);
            RunTwoTables(run, s);
            RunMixedOps(run, s);
            RunMisc(run, s);
        }

        private static void RunColspan(TestRun run, WdSession s)
        {
            string two = Tbl(W2, Tr(Td("合列头", 2)), Tr(Td(""), Td("")));
            string three = Tbl(W3, Tr(Td("合列头", 3)), Tr(Td(""), Td(""), Td("")));
            string one = Tbl(W2, Tr(Td("合列头", 2)), Tr(Td("有字"), Td("")));

            GeoN(run, s, "wm-c-col-two", two, 2, p => Rep(p, 2), GeoRun(2));
            GeoN(run, s, "wm-c-col-one", one, 1, p => Rep(p, 1), SingleGeo(1));
            GeoN(run, s, "wm-c-col-insrep", two, 2, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoInsert(p.Geos[0], A),
                WmCases.GeoReplace(p.Geos[1], B)
            }, GeoRun(2, "write via=map type=insert", "write via=map type=replace"));
            GeoN(run, s, "wm-c-col-del", two, 2, p => Del(p, 2), GeoRun(2, "write via=map type=delete"));
            GeoN(run, s, "wm-c-col-3", three, 3, p => Rep(p, 3), GeoRun(3));
            GeoN(run, s, "wm-c-col-ird", three, 3, p => Ird(p), GeoRun(3,
                "write via=map type=insert", "write via=map type=replace", "write via=map type=delete"));
            GeoN(run, s, "wm-c-col-flash", two, 2, p => Rep(p, 2), GeoRun(2),
                extraCounts: new[] { WmCases.Count("flash row=", 1) });
            GeoN(run, s, "wm-c-col-head", two, 2, p =>
            {
                var a = Rep(p, 2);
                if (!string.IsNullOrEmpty(p.HeadS))
                {
                    a.Add(WmCases.SentReplace(p.HeadS, "合列头改。"));
                }

                return a;
            }, Mixed(1, 1));
        }

        private static void RunRowspan(TestRun run, WdSession s)
        {
            string two = Tbl(W2, Tr(Td("合行左", 1, 2), Td("")), Tr(Td("")));
            string three = Tbl(W2,
                Tr(Td("合行左", 1, 2), Td("")),
                Tr(Td("")),
                Tr(Td(""), Td("")));
            string one = Tbl(W2, Tr(Td("合行左", 1, 2), Td("有")), Tr(Td("")));

            GeoN(run, s, "wm-c-row-two", two, 2, p => Rep(p, 2), GeoRun(2));
            GeoN(run, s, "wm-c-row-one", one, 1, p => Rep(p, 1), SingleGeo(1));
            GeoN(run, s, "wm-c-row-insrep", two, 2, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoInsert(p.Geos[0], A),
                WmCases.GeoReplace(p.Geos[1], B)
            }, GeoRun(2, "write via=map type=insert", "write via=map type=replace"));
            GeoN(run, s, "wm-c-row-del", two, 2, p => Del(p, 2), GeoRun(2, "write via=map type=delete"));
            GeoN(run, s, "wm-c-row-ird", three, 3, p => Ird(p), GeoRun(3,
                "write via=map type=insert", "write via=map type=replace", "write via=map type=delete"));
            GeoN(run, s, "wm-c-row-mix-del", two, 2, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoReplace(p.Geos[0], A),
                WmCases.GeoDelete(p.Geos[1])
            }, GeoRun(2, "write via=map type=replace", "write via=map type=delete"));
            GeoN(run, s, "wm-c-row-flash", two, 2, p => Rep(p, 2), GeoRun(2),
                extraCounts: new[] { WmCases.Count("flash row=", 2) });
            GeoN(run, s, "wm-c-row-order", two, 2, p => Rep(p, 2), null, custom: (id, wr, p) => OrderAssert(run, id, wr, p));
        }

        private static void RunRowCol(TestRun run, WdSession s)
        {
            // colspan 邻格不能写空 td：建表会当成占位重叠。余列用有字格，空格放在 rowspan 下方。
            string two = Tbl(W3,
                Tr(Td("行列块", 2, 2), Td("右")),
                Tr(Td("")),
                Tr(Td(""), Td("占"), Td("占")));
            string many = Tbl(W3,
                Tr(Td("行列块", 2, 2), Td("右")),
                Tr(Td("")),
                Tr(Td(""), Td(""), Td("")));

            GeoN(run, s, "wm-c-rc-two", two, 2, p => Rep(p, 2), GeoRun(2));
            GeoN(run, s, "wm-c-rc-insrep", two, 2, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoInsert(p.Geos[0], A),
                WmCases.GeoReplace(p.Geos[1], B)
            }, GeoRun(2, "write via=map type=insert", "write via=map type=replace"));
            GeoN(run, s, "wm-c-rc-del", two, 2, p => Del(p, 2), GeoRun(2, "write via=map type=delete"));
            GeoN(run, s, "wm-c-rc-3", many, 3, p => Rep(p, 3), GeoRun(3));
            GeoN(run, s, "wm-c-rc-ird", many, 3, p => Ird(p), GeoRun(3,
                "write via=map type=insert", "write via=map type=replace", "write via=map type=delete"));
            GeoN(run, s, "wm-c-rc-flash", two, 2, p => Rep(p, 2), GeoRun(2),
                extraCounts: new[] { WmCases.Count("flash row=", 2) });
        }

        private static void RunOutside(TestRun run, WdSession s)
        {
            string grid = Tbl(W2, Tr(Td("题外"), Td("")), Tr(Td("备外"), Td("")));
            MixOut(run, s, "wm-c-out-start", grid, 2, p =>
            {
                var a = Rep(p, 2);
                a.Add(InsStart(Outside));
                return a;
            }, Mixed(1, 1));
            MixOut(run, s, "wm-c-out-tbl", grid, 2, p =>
            {
                var a = Rep(p, 2);
                a.Add(InsAfterTable(p.TableId, After));
                return a;
            }, Mixed(1, 1));
            MixOut(run, s, "wm-c-out-all", grid, 2, p =>
            {
                var a = new List<Dictionary<string, object>>
                {
                    WmCases.GeoInsert(p.Geos[0], A),
                    WmCases.GeoReplace(p.Geos[1], B),
                    InsStart(Outside)
                };
                return a;
            }, Mixed(1, 1, "write via=map type=insert", "write via=map type=replace"));
            MixOut(run, s, "wm-c-out-delgeo", grid, 2, p =>
            {
                var a = Del(p, 2);
                a.Add(InsStart(Outside));
                return a;
            }, Mixed(1, 1, "write via=map type=delete"));
            MixOut(run, s, "wm-c-out-start-one", grid, 1, p =>
            {
                var a = Rep(p, 1);
                a.Add(InsStart(Outside));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=0", "singles=2", "step=single", "locate via=single" },
                new[] { "step=geo_run", "locate via=geo_run", "write via=map" },
                new[] { WmCases.Count("step=single", 2) }));
            MixOut(run, s, "wm-c-out-ird", grid, 2, p =>
            {
                return new List<Dictionary<string, object>>
                {
                    WmCases.GeoInsert(p.Geos[0], A),
                    WmCases.GeoDelete(p.Geos[1]),
                    InsAfterTable(p.TableId, After)
                };
            }, Mixed(1, 1, "write via=map type=insert", "write via=map type=delete"));

            PreOutside(run, s, "wm-c-out-rep", grid, 2, (p, outside) =>
            {
                var a = Rep(p, 2);
                a.Add(WmCases.SentReplace(outside, "表前句改。"));
                return a;
            }, Mixed(1, 1));
            PreOutside(run, s, "wm-c-out-del", grid, 2, (p, outside) =>
            {
                var a = Rep(p, 2);
                a.Add(SentDelete(outside));
                return a;
            }, Mixed(1, 1));
            PreOutside(run, s, "wm-c-out-ins-s", grid, 2, (p, outside) =>
            {
                var a = Rep(p, 2);
                a.Add(InsAfterCodes(outside, "续写。"));
                return a;
            }, Mixed(1, 1));
            PreOutside(run, s, "wm-c-out-note", Tbl(W3, Tr(Td("题外"), Td(""), Td(""))), 2, (p, outside) =>
            {
                var a = new List<Dictionary<string, object>>
                {
                    WmCases.GeoInsert(p.Geos[0], A),
                    WmCases.GeoReplace(p.Geos[1], B),
                    WmCases.SentReplace(outside, "表前句改。")
                };
                if (!string.IsNullOrEmpty(p.HeadS))
                {
                    a.Add(WmCases.SentReplace(p.HeadS, "题外改。"));
                }

                return a;
            }, Mixed(1, 2));
        }

        private static void RunTwoTables(TestRun run, WdSession s)
        {
            TwoT(run, s, "wm-c-2t-each", (p1, p2) =>
            {
                var a = Rep(p1, 2);
                a.AddRange(Rep(p2, 2));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=2", "singles=0", "step=geo_run", "locate via=geo_run" },
                new[] { "locate via=single", "step=single" },
                new[]
                {
                    WmCases.Count("step=geo_run", 2),
                    WmCases.Count("locate via=geo_run", 2),
                    WmCases.Count("write via=map", 4),
                    WmCases.Count("post_pd mode=geo_run", 2),
                    WmCases.Count("map=once", 2)
                }));
            TwoT(run, s, "wm-c-2t-one-each", (p1, p2) =>
            {
                var a = Rep(p1, 1);
                a.AddRange(Rep(p2, 1));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=0", "singles=2", "step=single", "locate via=single" },
                new[] { "step=geo_run", "locate via=geo_run", "write via=map" },
                new[] { WmCases.Count("step=single", 2), WmCases.Count("post_pd mode=single", 2) }));
            TwoT(run, s, "wm-c-2t-mix", (p1, p2) =>
            {
                var a = Rep(p1, 2);
                a.AddRange(Del(p2, 2));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=2", "write via=map type=replace", "write via=map type=delete" },
                new[] { "locate via=single" },
                new[] { WmCases.Count("step=geo_run", 2), WmCases.Count("write via=map", 4) }));
            TwoT(run, s, "wm-c-2t-split-s", (p1, p2) =>
            {
                var a = Rep(p1, 2);
                a.Add(InsAfterTable(p1.TableId, After));
                a.AddRange(Rep(p2, 2));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=2", "singles=1", "step=single", "step=geo_run" },
                new[] { "locate via=single" },
                new[]
                {
                    WmCases.Count("step=geo_run", 2),
                    WmCases.Count("step=single", 1),
                    WmCases.Count("write via=map", 4)
                }));
            TwoT(run, s, "wm-c-2t-ird", (p1, p2) =>
            {
                var a = new List<Dictionary<string, object>>
                {
                    WmCases.GeoInsert(p1.Geos[0], A),
                    WmCases.GeoReplace(p1.Geos[1], B),
                    WmCases.GeoDelete(p2.Geos[0]),
                    WmCases.GeoReplace(p2.Geos[1], C)
                };
                return a;
            }, new Expect(
                new[]
                {
                    "plan geo_runs=2",
                    "write via=map type=insert",
                    "write via=map type=replace",
                    "write via=map type=delete"
                },
                new[] { "locate via=single" },
                new[] { WmCases.Count("step=geo_run", 2), WmCases.Count("write via=map", 4) }));
            TwoT(run, s, "wm-c-2t-out", (p1, p2) =>
            {
                var a = Rep(p1, 2);
                a.AddRange(Rep(p2, 2));
                a.Add(InsStart(Outside));
                return a;
            }, new Expect(
                new[] { "plan geo_runs=2", "singles=1", "step=single", "step=geo_run" },
                new[] { "locate via=single" },
                new[] { WmCases.Count("step=geo_run", 2), WmCases.Count("step=single", 1) }));
        }

        private static void RunMixedOps(TestRun run, WdSession s)
        {
            string g23 = Tbl(W3, Tr(Td("题三"), Td(""), Td("")), Tr(Td(""), Td(""), Td("")));
            GeoN(run, s, "wm-c-ird", g23, 3, p => Ird(p), GeoRun(3,
                "write via=map type=insert", "write via=map type=replace", "write via=map type=delete"));
            GeoN(run, s, "wm-c-rri", g23, 3, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoReplace(p.Geos[0], A),
                WmCases.GeoReplace(p.Geos[1], B),
                WmCases.GeoInsert(p.Geos[2], C)
            }, GeoRun(3, "write via=map type=replace", "write via=map type=insert"));
            GeoN(run, s, "wm-c-ddi", g23, 3, p => new List<Dictionary<string, object>>
            {
                WmCases.GeoDelete(p.Geos[0]),
                WmCases.GeoDelete(p.Geos[1]),
                WmCases.GeoInsert(p.Geos[2], C)
            }, GeoRun(3, "write via=map type=delete", "write via=map type=insert"));
            GeoN(run, s, "wm-c-title-geo", g23, 2, p =>
            {
                var a = Rep(p, 2);
                if (!string.IsNullOrEmpty(p.HeadS))
                {
                    a.Add(WmCases.SentReplace(p.HeadS, "题三改。"));
                }

                return a;
            }, Mixed(1, 1));
            GeoN(run, s, "wm-c-ins3", g23, 3, p => Ins(p, 3), GeoRun(3, "write via=map type=insert"));
            GeoN(run, s, "wm-c-rep3", g23, 3, p => Rep(p, 3), GeoRun(3));
            GeoN(run, s, "wm-c-del3", g23, 3, p => Del(p, 3), GeoRun(3, "write via=map type=delete"));
            string row3 = Tbl(W3, Tr(Td(""), Td(""), Td("")));
            GeoN(run, s, "wm-c-same-row-3", row3, 3, p => Rep(p, 3), GeoRun(3),
                extraCounts: new[] { WmCases.Count("flash row=", 1) });
        }

        private static void RunMisc(TestRun run, WdSession s)
        {
            string two = Tbl(W2, Tr(Td("杂题"), Td("")), Tr(Td("杂备"), Td("")));
            GeoN(run, s, "wm-c-bold", two, 2, p => new List<Dictionary<string, object>>
            {
                GeoReplaceFmt(p.Geos[0], A, true),
                GeoReplaceFmt(p.Geos[1], B, false)
            }, GeoRun(2));
            GeoN(run, s, "wm-c-runs", two, 2, p => new List<Dictionary<string, object>>
            {
                GeoInsertRuns(p.Geos[0]),
                WmCases.GeoReplace(p.Geos[1], B)
            }, GeoRun(2, "write via=map type=insert", "write via=map type=replace"));
            GeoN(run, s, "wm-c-echo3", Tbl(W3, Tr(Td(""), Td(""), Td(""))), 3, p => Rep(p, 3), null, custom:
                (caseId, wr, _) =>
                {
                    string blob = WmCases.Blob(wr);
                    bool ok = wr != null && wr.Success
                        && WdSession.HasKey(wr, "geo_plan_log")
                        && blob.IndexOf("step=geo_run cells=3", StringComparison.Ordinal) >= 0
                        && blob.IndexOf(DocumentActionPathLog.LogTag, StringComparison.Ordinal) >= 0;
                    WmCases.Ok(run, caseId, ok, wr == null || !wr.Success
                        ? (wr == null ? "null" : wr.Error)
                        : WmCases.TrimBlob(blob));
                });
            GeoN(run, s, "wm-c-p-mix", two, 2, p =>
            {
                var a = Rep(p, 2);
                a.Add(ParaReplace(p.Para, "段改。\r"));
                return a;
            }, Mixed(1, 1), wantPara: true);
        }

        private sealed class Plant
        {
            public string Path;
            public string Id;
            public string TableId;
            public string HeadS;
            public string Para;
            public List<string> Geos = new List<string>();
        }

        private sealed class Expect
        {
            public string[] Have;
            public string[] Forbid;
            public KeyValuePair<string, int>[] Counts;

            public Expect(string[] have, string[] forbid, KeyValuePair<string, int>[] counts)
            {
                Have = have;
                Forbid = forbid;
                Counts = counts;
            }
        }

        private static Expect GeoRun(int cells, params string[] extraHave)
        {
            var have = new List<string>
            {
                "plan geo_runs=1",
                "step=geo_run cells=" + cells,
                "locate via=geo_run n=1",
                "map=once",
                "write via=map",
                "post_pd mode=geo_run"
            };
            have.AddRange(extraHave);
            return new Expect(
                have.ToArray(),
                new[] { "locate via=single", "step=single", "post_pd mode=single" },
                new[]
                {
                    WmCases.Count("write via=map", cells),
                    WmCases.Count("locate via=geo_run", 1),
                    WmCases.Count("post_pd mode=geo_run", 1)
                });
        }

        private static Expect SingleGeo(int n)
        {
            return new Expect(
                new[] { "plan geo_runs=0", "singles=" + n, "step=single", "geo=true", "locate via=single" },
                new[] { "step=geo_run", "locate via=geo_run", "write via=map", "post_pd mode=geo_run" },
                new[] { WmCases.Count("step=single", n), WmCases.Count("locate via=single", n) });
        }

        private static Expect Mixed(int geoRuns, int singles, params string[] extraHave)
        {
            var have = new List<string>
            {
                "plan geo_runs=" + geoRuns,
                "singles=" + singles,
                "step=geo_run",
                "step=single",
                "locate via=geo_run"
            };
            have.AddRange(extraHave);
            return new Expect(
                have.ToArray(),
                new[] { "locate via=single" },
                new[]
                {
                    WmCases.Count("step=geo_run", geoRuns),
                    WmCases.Count("step=single", singles),
                    WmCases.Count("locate via=geo_run", geoRuns)
                });
        }

        private static void GeoN(
            TestRun run,
            WdSession s,
            string id,
            string xml,
            int need,
            Func<Plant, List<Dictionary<string, object>>> acts,
            Expect expect,
            KeyValuePair<string, int>[] extraCounts = null,
            Action<string, ToolResult, Plant> custom = null,
            bool wantPara = false)
        {
            Plant p;
            if (!TryPlant(run, s, xml, need, id, out p))
            {
                return;
            }

            if (wantPara)
            {
                ToolResult paraGot = s.GetContent(p.Id, "paragraph");
                string paraBody = paraGot != null && paraGot.Success
                    ? WdSession.Str(paraGot, "document_content")
                    : "";
                p.Para = FirstP(paraBody);
                if (string.IsNullOrEmpty(p.Para))
                {
                    run.CaseFail(id, "无 P_ " + WmCases.TrimBlob(paraBody));
                    s.CloseDocByPath(p.Path);
                    return;
                }
            }

            ToolResult wr = s.ProcessActions(p.Id, acts(p));
            Finish(run, id, wr, p, expect, extraCounts, custom);
            s.CloseDocByPath(p.Path);
        }

        private static void MixOut(
            TestRun run,
            WdSession s,
            string id,
            string xml,
            int need,
            Func<Plant, List<Dictionary<string, object>>> acts,
            Expect expect)
        {
            GeoN(run, s, id, xml, need, acts, expect);
        }

        private static void PreOutside(
            TestRun run,
            WdSession s,
            string id,
            string xml,
            int need,
            Func<Plant, string, List<Dictionary<string, object>>> acts,
            Expect expect)
        {
            Plant p;
            if (!TryPlant(run, s, xml, need, id, out p))
            {
                return;
            }

            ToolResult pre = s.ProcessActions(p.Id, new List<Dictionary<string, object>> { InsStart(Outside) });
            if (pre == null || !pre.Success)
            {
                run.CaseFail(id, pre == null ? "预插表前句失败" : (pre.Error ?? "预插失败"));
                s.CloseDocByPath(p.Path);
                return;
            }

            ToolResult got = s.GetContent(p.Id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            string outside = FindS(body, Outside);
            if (string.IsNullOrEmpty(outside))
            {
                run.CaseFail(id, "无表前句 S_");
                s.CloseDocByPath(p.Path);
                return;
            }

            string titleKeep = p.HeadS;
            FillCodes(p, body);
            if (!string.IsNullOrEmpty(titleKeep))
            {
                p.HeadS = titleKeep;
            }

            ToolResult wr = s.ProcessActions(p.Id, acts(p, outside));
            Finish(run, id, wr, p, expect, null, null);
            s.CloseDocByPath(p.Path);
        }

        private static void TwoT(
            TestRun run,
            WdSession s,
            string id,
            Func<Plant, Plant, List<Dictionary<string, object>>> acts,
            Expect expect)
        {
            string xml = Tbl(W2, Tr(Td("两表题"), Td("")), Tr(Td("两表备"), Td("")));
            Plant p1;
            if (!TryPlant(run, s, xml, 2, id, out p1))
            {
                return;
            }

            var args = WmCases.CreateArgs(p1.Id, WmCases.WriteXml(xml));
            args["target"] = new Dictionary<string, object> { ["table"] = p1.TableId };
            ToolResult second = s.CreateTableFromXml(args);
            string t2 = WdSession.Str(second, "table_id");
            if (!second.Success || string.IsNullOrEmpty(t2))
            {
                run.CaseFail(id, second.Error ?? "第二表失败");
                s.CloseDocByPath(p1.Path);
                return;
            }

            ToolResult got = s.GetContent(p1.Id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            var all = CollectGeos(body);
            var g1 = new List<string>();
            var g2 = new List<string>();
            foreach (string g in all)
            {
                EmptyCellCode.Parsed parsed;
                if (!EmptyCellCode.TryParse(g, out parsed))
                {
                    continue;
                }

                if (parsed.TableId == p1.TableId)
                {
                    g1.Add(g);
                }
                else if (parsed.TableId == t2)
                {
                    g2.Add(g);
                }
            }

            if (g1.Count < 2 || g2.Count < 2)
            {
                run.CaseFail(id, "两表空格不足 g1=" + g1.Count + " g2=" + g2.Count);
                s.CloseDocByPath(p1.Path);
                return;
            }

            var p2 = new Plant { Path = p1.Path, Id = p1.Id, TableId = t2, Geos = g2 };
            p1.Geos = g1;
            ToolResult wr = s.ProcessActions(p1.Id, acts(p1, p2));
            Finish(run, id, wr, p1, expect, null, null);
            s.CloseDocByPath(p1.Path);
        }

        private static void Finish(
            TestRun run,
            string id,
            ToolResult wr,
            Plant p,
            Expect expect,
            KeyValuePair<string, int>[] extraCounts,
            Action<string, ToolResult, Plant> custom)
        {
            if (custom != null)
            {
                custom(id, wr, p);
                return;
            }

            if (wr != null && !wr.Success)
            {
                run.CaseFail(id, wr.Error ?? "失败 " + WmCases.TrimBlob(WmCases.Blob(wr)));
                return;
            }

            var counts = new List<KeyValuePair<string, int>>();
            if (expect.Counts != null)
            {
                counts.AddRange(expect.Counts);
            }

            if (extraCounts != null)
            {
                counts.AddRange(extraCounts);
            }

            WmCases.ExpectPath(run, id, wr, expect.Have, expect.Forbid, counts.ToArray());
        }

        private static void OrderAssert(TestRun run, string id, ToolResult wr, Plant p)
        {
            EmptyCellCode.Parsed pa;
            EmptyCellCode.Parsed pb;
            if (p.Geos.Count < 2
                || !EmptyCellCode.TryParse(p.Geos[0], out pa)
                || !EmptyCellCode.TryParse(p.Geos[1], out pb))
            {
                run.CaseFail(id, "空格码解析失败");
                return;
            }

            bool aLater = pa.Row > pb.Row || (pa.Row == pb.Row && pa.Col > pb.Col);
            int laterR = aLater ? pa.Row : pb.Row;
            int laterC = aLater ? pa.Col : pb.Col;
            int earlierR = aLater ? pb.Row : pa.Row;
            int earlierC = aLater ? pb.Col : pa.Col;
            string blob = WmCases.Blob(wr);
            string laterTok = "write via=map type=replace r=" + laterR + " c=" + laterC;
            string earlierTok = "write via=map type=replace r=" + earlierR + " c=" + earlierC;
            int iLater = blob.IndexOf(laterTok, StringComparison.Ordinal);
            int iEarlier = blob.IndexOf(earlierTok, StringComparison.Ordinal);
            bool geo = blob.IndexOf("step=geo_run cells=2", StringComparison.Ordinal) >= 0
                && blob.IndexOf("locate via=single", StringComparison.Ordinal) < 0;
            WmCases.Ok(run, id,
                wr != null && wr.Success && geo && iLater >= 0 && iEarlier >= 0 && iLater < iEarlier,
                wr == null || !wr.Success
                    ? (wr == null ? "null" : wr.Error)
                    : ("later@" + iLater + " earlier@" + iEarlier + " " + WmCases.TrimBlob(blob)));
        }

        private static bool TryPlant(TestRun run, WdSession s, string xml, int need, string failId, out Plant plant)
        {
            plant = new Plant();
            plant.Path = s.NewDocxPath();
            ToolResult opened = s.Open(plant.Path, true, s.App);
            if (!opened.Success)
            {
                WmCases.FailMany(run, opened.Error ?? "新建失败", failId);
                return false;
            }

            plant.Id = WdSession.Str(opened, "channel_id");
            s.ActivateDoc(plant.Path);
            ToolResult created = s.CreateTableFromXml(WmCases.CreateArgs(plant.Id, WmCases.WriteXml(xml)));
            plant.TableId = WdSession.Str(created, "table_id");
            if (!created.Success || string.IsNullOrEmpty(plant.TableId))
            {
                WmCases.FailMany(run, created.Error ?? "建表失败", failId);
                s.CloseDocByPath(plant.Path);
                return false;
            }

            ToolResult got = s.GetContent(plant.Id);
            string body = got.Success ? WdSession.Str(got, "document_content") : "";
            FillCodes(plant, body);
            if (plant.Geos.Count < need)
            {
                WmCases.FailMany(run, "空格码不足 need=" + need + " got=" + plant.Geos.Count + " " + WmCases.TrimBlob(body), failId);
                s.CloseDocByPath(plant.Path);
                return false;
            }

            return true;
        }

        private static void FillCodes(Plant plant, string body)
        {
            plant.Geos = CollectGeos(body);
            plant.HeadS = FirstS(body);
            plant.Para = FirstP(body);
        }

        private static List<string> CollectGeos(string body)
        {
            var list = new List<string>();
            foreach (Match m in Regex.Matches(body ?? "", @"\[(T_[0-9A-Za-z]+_r[0-9]+_c[0-9]+)-start\]"))
            {
                string code = m.Groups[1].Value;
                if (!list.Contains(code))
                {
                    list.Add(code);
                }
            }

            return list;
        }

        private static string FirstS(string body)
        {
            Match m = Regex.Match(body ?? "", @"\[(S_[0-9a-zA-Z]{5})-start");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string FindS(string body, string text)
        {
            Match m = Regex.Match(
                body ?? "",
                @"\[(S_[0-9a-zA-Z]{5})-start(?: [^\]]*)?\](?:\[in_table=[^\]]+\])?" + Regex.Escape(text ?? ""));
            return m.Success ? m.Groups[1].Value : "";
        }

        private static string FirstP(string body)
        {
            Match m = Regex.Match(body ?? "", @"\[(P_[0-9a-zA-Z]+(?:_[0-9]+)?)(?:-start| start)");
            return m.Success ? m.Groups[1].Value : "";
        }

        private static List<Dictionary<string, object>> Rep(Plant p, int n)
        {
            var a = new List<Dictionary<string, object>>();
            string[] texts = { A, B, C };
            for (int i = 0; i < n; i++)
            {
                a.Add(WmCases.GeoReplace(p.Geos[i], texts[i % texts.Length]));
            }

            return a;
        }

        private static List<Dictionary<string, object>> Ins(Plant p, int n)
        {
            var a = new List<Dictionary<string, object>>();
            string[] texts = { A, B, C };
            for (int i = 0; i < n; i++)
            {
                a.Add(WmCases.GeoInsert(p.Geos[i], texts[i % texts.Length]));
            }

            return a;
        }

        private static List<Dictionary<string, object>> Del(Plant p, int n)
        {
            var a = new List<Dictionary<string, object>>();
            for (int i = 0; i < n; i++)
            {
                a.Add(WmCases.GeoDelete(p.Geos[i]));
            }

            return a;
        }

        private static List<Dictionary<string, object>> Ird(Plant p)
        {
            return new List<Dictionary<string, object>>
            {
                WmCases.GeoInsert(p.Geos[0], A),
                WmCases.GeoReplace(p.Geos[1], B),
                WmCases.GeoDelete(p.Geos[2])
            };
        }

        private static Dictionary<string, object> InsStart(string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["position"] = "start" },
                    ["insert"] = text,
                    ["format"] = WmCases.DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> InsAfterTable(string tableId, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["table"] = tableId },
                    ["insert"] = text,
                    ["format"] = WmCases.DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> InsAfterCodes(string codes, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["codes"] = codes },
                    ["insert"] = text,
                    ["format"] = WmCases.DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> SentDelete(string code)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "delete",
                ["detail"] = new Dictionary<string, object>
                {
                    ["deletion"] = new Dictionary<string, object> { ["codes"] = code }
                }
            };
        }

        private static Dictionary<string, object> ParaReplace(string code, string text)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["paragraph_codes"] = code },
                    ["new"] = text,
                    ["format"] = WmCases.DefaultFormat()
                }
            };
        }

        private static Dictionary<string, object> GeoReplaceFmt(string code, string text, bool bold)
        {
            return new Dictionary<string, object>
            {
                ["type"] = "replace",
                ["detail"] = new Dictionary<string, object>
                {
                    ["original"] = new Dictionary<string, object> { ["codes"] = code },
                    ["new"] = text,
                    ["format"] = new Dictionary<string, object>
                    {
                        ["char_format"] = new Dictionary<string, object>
                        {
                            ["font_name"] = "宋体",
                            ["font_size"] = 12,
                            ["is_bold"] = bold
                        }
                    }
                }
            };
        }

        private static Dictionary<string, object> GeoInsertRuns(string code)
        {
            var runs = new List<object>
            {
                new Dictionary<string, object>
                {
                    ["text"] = "片一。",
                    ["format"] = WmCases.DefaultFormat()
                },
                new Dictionary<string, object>
                {
                    ["text"] = "片二。",
                    ["format"] = WmCases.DefaultFormat()
                }
            };
            return new Dictionary<string, object>
            {
                ["type"] = "insert",
                ["detail"] = new Dictionary<string, object>
                {
                    ["target"] = new Dictionary<string, object> { ["codes"] = code },
                    ["insert"] = runs
                }
            };
        }

        private const string W2 = "<float>120</float><float>180</float>";
        private const string W3 = "<float>100</float><float>100</float><float>100</float>";

        private static string Tbl(string widths, params string[] rows)
        {
            return
                "<TableConfig><General><TableFontColor>000000</TableFontColor></General><Table><Properties><ColWidths>"
                + widths + "</ColWidths></Properties><Data>"
                + string.Concat(rows)
                + "</Data></Table></TableConfig>";
        }

        private static string Tr(params string[] tds)
        {
            return "<tr>" + string.Concat(tds) + "</tr>";
        }

        private static string Td(string text, int colspan = 1, int rowspan = 1)
        {
            string a = "";
            if (colspan > 1)
            {
                a += " colspan=\"" + colspan + "\"";
            }

            if (rowspan > 1)
            {
                a += " rowspan=\"" + rowspan + "\"";
            }

            return "<td" + a + ">" + (text ?? "") + "</td>";
        }
    }
}
