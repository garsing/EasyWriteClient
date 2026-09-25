using System;
using System.Collections.Generic;
using WordAddIn1.PresentationHost;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>统一 OOXML 改包+瘦包贴回。前缀 oox-，30 条 × 两宿主。</summary>
    internal static partial class ContentCatalog
    {
        private enum OoxPos
        {
            First,
            Mid,
            Last
        }

        private static readonly string[] OoxCats = { "A", "B", "C" };
        private static readonly string[] OoxVals = { "40", "35", "25" };
        private static readonly string[] OoxCats2 = { "X", "Y", "Z" };
        private static readonly string[] OoxVals2 = { "10", "20", "70" };

        private static void AddOoxmlCases(List<ContentCase> list, ref int page)
        {
            AddOox(list, ref page, "oox-create-pie-mid",
                s => RunCreate(s, "pie2d", OoxPos.Mid, false, false));
            AddOox(list, ref page, "oox-create-col-mid",
                s => RunCreate(s, "column", OoxPos.Mid, false, false));
            AddOox(list, ref page, "oox-create-bar-mid",
                s => RunCreate(s, "bar", OoxPos.Mid, false, false));
            AddOox(list, ref page, "oox-create-line-mid",
                s => RunCreate(s, "line", OoxPos.Mid, false, false));
            AddOox(list, ref page, "oox-create-pie-last",
                s => RunCreate(s, "pie2d", OoxPos.Last, false, false));
            AddOox(list, ref page, "oox-create-tb-pie-mid",
                s => RunCreate(s, "pie2d", OoxPos.Mid, true, false));
            AddOox(list, ref page, "oox-create-two-pie-mid",
                s => RunCreate(s, "pie2d", OoxPos.Mid, false, true));

            AddOox(list, ref page, "oox-lean-pie-mid",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-col-mid",
                s => RunLean(s, "column", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-bar-mid",
                s => RunLean(s, "bar", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-line-mid",
                s => RunLean(s, "line", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-pie-last",
                s => RunLean(s, "pie2d", OoxPos.Last, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-pie-first",
                s => RunLean(s, "pie2d", OoxPos.First, LeanFlag.None));
            AddOox(list, ref page, "oox-lean-two-pie-one",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.TwoCharts));
            AddOox(list, ref page, "oox-lean-pie-keep-geo",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.KeepGeo));
            AddOox(list, ref page, "oox-lean-col-keep-title",
                s => RunLean(s, "column", OoxPos.Mid, LeanFlag.KeepTitle));
            AddOox(list, ref page, "oox-lean-pie-with-tb",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.KeepTitle));

            AddOox(list, ref page, "oox-grp-nested-mid",
                s => RunNested(s, OoxPos.Mid, false));
            AddOox(list, ref page, "oox-grp-nested-last",
                s => RunNested(s, OoxPos.Last, false));
            AddOox(list, ref page, "oox-grp-nested-first",
                s => RunNested(s, OoxPos.First, false));
            AddOox(list, ref page, "oox-grp-three-tb",
                s => RunThreeWrap(s));
            AddOox(list, ref page, "oox-grp-nested-sib",
                s => RunNested(s, OoxPos.Mid, true));

            AddOox(list, ref page, "oox-pipe-slim-log",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.RequireSlim));
            AddOox(list, ref page, "oox-pipe-dest-pages",
                s => RunDestPages(s));
            AddOox(list, ref page, "oox-pipe-dest-alive",
                s => RunDestAlive(s));
            AddOox(list, ref page, "oox-pipe-no-temp",
                s => RunLean(s, "column", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-pipe-sib-keep",
                s => RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.None));
            AddOox(list, ref page, "oox-pipe-pres-count",
                s => RunPresCount(s));

            AddOox(list, ref page, "oox-read-chart-grid",
                s => RunReadGrid(s));
            AddOox(list, ref page, "oox-neg-bad-chart-id",
                s => RunNegBadId(s));
        }

        [Flags]
        private enum LeanFlag
        {
            None = 0,
            TwoCharts = 1,
            KeepGeo = 2,
            KeepTitle = 4,
            RequireSlim = 8
        }

        private static string RunCreate(
            ContentGroupSession s,
            string chartType,
            OoxPos pos,
            bool withTitle,
            bool twoCharts)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            int startPres = s.PresCount();
            OoxDeck deck;
            string err = SeedDeck(s, pos, out deck);
            if (err != null)
            {
                return err;
            }

            string html = twoCharts
                ? ContentHtml.MixParts(
                    ContentHtml.ChartFrag("pie2d", 6, 18, 40, 50, OoxCats, OoxVals),
                    ContentHtml.ChartFrag("pie2d", 52, 18, 40, 50, OoxCats2, OoxVals2))
                : withTitle
                    ? ContentHtml.MixParts(
                        ContentHtml.MixTitleMarkup("Oox标题"),
                        ContentHtml.ChartFrag(chartType, 8, 22, 50, 50, OoxCats, OoxVals))
                    : ContentHtml.ChartCreate(chartType, 8, 18, 50, 55, OoxCats, OoxVals);

            if (!s.TryApply(html, true, out List<string> ids, out err) || ids.Count < 1)
            {
                return "建图失败: " + err + DumpWarn(s);
            }

            string pipe = AssertPipe(s, deck, temp, startPres, requireSlim: true);
            if (pipe != null)
            {
                return pipe;
            }

            if (!s.TryReadFullPage(out PptHtmlReadResult page, out err))
            {
                return "建图后读失败: " + err;
            }

            if (withTitle)
            {
                string te = TextHas(page, "Oox标题");
                if (te != null)
                {
                    return te;
                }
            }

            List<PptHtmlShapeNode> charts = FindAll(page, "chart");
            if (twoCharts)
            {
                if (charts.Count < 2)
                {
                    return "双饼只读到 " + charts.Count;
                }

                return ContentAssert.ChartDataMismatch(charts[0], OoxCats, OoxVals, "pie")
                    ?? ContentAssert.ChartDataMismatch(charts[1], OoxCats2, OoxVals2, "pie");
            }

            if (charts.Count < 1)
            {
                return "建图后没有 chart";
            }

            return ContentAssert.ChartDataMismatch(charts[0], OoxCats, OoxVals, chartType);
        }

        private static string RunLean(ContentGroupSession s, string chartType, OoxPos pos, LeanFlag flags)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            int startPres = s.PresCount();
            OoxDeck deck;
            string err = SeedDeck(s, pos, out deck);
            if (err != null)
            {
                return err;
            }

            bool two = (flags & LeanFlag.TwoCharts) != 0;
            bool keepTitle = (flags & LeanFlag.KeepTitle) != 0;
            string create = two
                ? ContentHtml.MixParts(
                    ContentHtml.ChartFrag("pie2d", 6, 18, 40, 50, OoxCats, OoxVals),
                    ContentHtml.ChartFrag("pie2d", 52, 18, 40, 50, OoxCats2, OoxVals2))
                : keepTitle
                    ? ContentHtml.MixParts(
                        ContentHtml.MixTitleMarkup("柱数原"),
                        ContentHtml.ChartFrag(chartType, 8, 22, 50, 50, OoxCats, OoxVals))
                    : ContentHtml.ChartCreate(chartType, 8, 18, 50, 55, OoxCats, OoxVals);

            if (!s.TryApply(create, true, out List<string> ids, out err) || ids.Count < 1)
            {
                return "铺底失败: " + err + DumpWarn(s);
            }

            if (!s.TryReadFullPage(out PptHtmlReadResult before, out err))
            {
                return "换数前读失败: " + err;
            }

            List<PptHtmlShapeNode> beforeCharts = FindAll(before, "chart");
            if (beforeCharts.Count < (two ? 2 : 1))
            {
                return "铺底缺图 " + beforeCharts.Count;
            }

            string geo = beforeCharts[two ? 1 : 0].Style;
            string chartId = beforeCharts[two ? 1 : 0].ShapeId;
            if (string.IsNullOrEmpty(chartId))
            {
                chartId = FindChartId(ids);
            }

            if (string.IsNullOrEmpty(chartId))
            {
                return "没有 chart ShapeId";
            }

            string seriesType = two ? "pie2d" : chartType;
            if (!s.TryApply(
                    ContentHtml.ChartUpdate(chartId, OoxCats2, OoxVals2, "规模", seriesType),
                    false,
                    out _,
                    out err))
            {
                return "换数失败: " + err + DumpWarn(s);
            }

            bool requireSlim = true;
            string pipe = AssertPipe(s, deck, temp, startPres, requireSlim);
            if (pipe != null)
            {
                return pipe;
            }

            if (!s.TryReadFullPage(out PptHtmlReadResult after, out err))
            {
                return "换数后读失败: " + err;
            }

            List<PptHtmlShapeNode> afterCharts = FindAll(after, "chart");
            if (afterCharts.Count < (two ? 2 : 1))
            {
                return "换数后缺图";
            }

            if (two)
            {
                string keep = ContentAssert.ChartDataMismatch(afterCharts[0], OoxCats, OoxVals, "pie");
                if (keep != null)
                {
                    return "误伤第一张饼: " + keep;
                }

                return ContentAssert.ChartDataMismatch(afterCharts[1], OoxCats2, OoxVals2, "pie");
            }

            string data = ContentAssert.ChartDataMismatch(afterCharts[0], OoxCats2, OoxVals2, chartType);
            if (data != null)
            {
                return data;
            }

            if (keepTitle)
            {
                string te = TextHas(after, "柱数原");
                if (te != null)
                {
                    return "换数误伤标题: " + te;
                }
            }

            if ((flags & LeanFlag.KeepGeo) != 0
                && !ContentAssert.GeoClose(geo, afterCharts[0].Style))
            {
                return "换数挪了框 前=" + geo + " 后=" + afterCharts[0].Style;
            }

            return null;
        }

        private static string RunNested(ContentGroupSession s, OoxPos pos, bool extraSibCheck)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            int startPres = s.PresCount();
            OoxDeck deck;
            string err = SeedDeck(s, pos, out deck);
            if (err != null)
            {
                return err;
            }

            if (!s.TryApply(ContentHtml.NestedGroups(), true, out _, out err))
            {
                return "嵌套组失败: " + err + DumpWarn(s);
            }

            string pipe = AssertPipe(s, deck, temp, startPres, requireSlim: false);
            if (pipe != null)
            {
                return pipe;
            }

            if (!s.TryReadPage(out PptHtmlReadResult page, out err))
            {
                return err;
            }

            PptHtmlShapeNode outer = ContentGroupSession.FindGroup(page);
            if (outer == null)
            {
                return "嵌套组后顶层无组";
            }

            if (!s.TryReadShape(outer.ShapeId, out PptHtmlReadResult expand, out err))
            {
                return "展开外组失败: " + err;
            }

            if (ContentGroupSession.FindChildGroup(
                    ContentGroupSession.FindByIdDeep(expand, outer.ShapeId) ?? outer) == null)
            {
                return "不见内组 树=" + ContentGroupSession.DescribeTree(expand);
            }

            return null;
        }

        private static string RunThreeWrap(ContentGroupSession s)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            int startPres = s.PresCount();
            OoxDeck deck;
            string err = SeedDeck(s, OoxPos.Mid, out deck);
            if (err != null)
            {
                return err;
            }

            if (!s.TryApply(ContentHtml.ThreeTextboxes(), true, out List<string> ids, out err)
                || ids.Count < 3)
            {
                return "铺三框失败: " + err;
            }

            if (!s.TryGroup(new[] { ids[0], ids[1] }, out PresentationManageShapeResult inner, out err)
                || string.IsNullOrEmpty(inner?.GroupShapeId))
            {
                return "内组失败: " + err;
            }

            if (!s.TryGroup(new[] { inner.GroupShapeId, ids[2] }, out PresentationManageShapeResult outer, out err)
                || string.IsNullOrEmpty(outer?.GroupShapeId))
            {
                return "外组(OOXML)失败: " + err;
            }

            string pipe = AssertPipe(s, deck, temp, startPres, requireSlim: false);
            if (pipe != null)
            {
                return pipe;
            }

            if (!s.TryReadPage(out PptHtmlReadResult page, out err))
            {
                return err;
            }

            return ContentGroupSession.FindGroup(page) == null ? "三框外组后顶层无组" : null;
        }

        private static string RunDestPages(ContentGroupSession s)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            int start = s.SlideCount();
            var extra = new List<string>();
            for (int i = 0; i < 4; i++)
            {
                string id;
                string err;
                if (!s.TryAddBlankSlide(out id, out err))
                {
                    return "加页失败: " + err;
                }

                extra.Add(id);
                if (!s.TrySelectSlide(id, out err)
                    || !s.TryApply(
                        ContentHtml.TextboxCreate("P" + i, left: 8, top: 8, width: 20, height: 8),
                        true,
                        out _,
                        out err))
                {
                    return "铺页失败: " + err;
                }
            }

            if (!s.TrySelectSlide(extra[1], out string goErr))
            {
                return goErr;
            }

            if (!s.TryApply(ContentHtml.ChartCreate("pie2d", 8, 18, 50, 55, OoxCats, OoxVals), true, out _, out string cErr))
            {
                return "5 页中建图失败: " + cErr;
            }

            if (s.SlideCount() != start + 4)
            {
                return "页数变了 期望 " + (start + 4) + " 实际 " + s.SlideCount();
            }

            return ContentGroupSession.LeftoverEwTemp(temp);
        }

        private static string RunDestAlive(ContentGroupSession s)
        {
            string lean = RunLean(s, "pie2d", OoxPos.Mid, LeanFlag.None);
            if (lean != null)
            {
                return lean;
            }

            if (!s.PresAlive())
            {
                return "换数后原稿死了";
            }

            if (!s.TryApply(
                    ContentHtml.TextboxCreate("ALIVE", left: 70, top: 80, width: 22, height: 8),
                    true,
                    out _,
                    out string err))
            {
                return "换数后再 apply 失败: " + err;
            }

            if (!s.TryReadPage(out PptHtmlReadResult page, out err))
            {
                return err;
            }

            return TextHas(page, "ALIVE");
        }

        private static string RunPresCount(ContentGroupSession s)
        {
            int before = s.PresCount();
            string lean = RunLean(s, "column", OoxPos.Mid, LeanFlag.None);
            if (lean != null)
            {
                return lean;
            }

            int after = s.PresCount();
            if (after != before)
            {
                return "稿数变了 前=" + before + " 后=" + after;
            }

            return s.PresAlive() ? null : "原稿死了";
        }

        private static string RunReadGrid(ContentGroupSession s)
        {
            int before = s.PresCount();
            if (!s.TryApply(
                    ContentHtml.ChartCreate("pie2d", 8, 18, 50, 55, OoxCats, OoxVals),
                    true,
                    out _,
                    out string err))
            {
                return "建图失败: " + err;
            }

            if (s.PresCount() != before)
            {
                return "建图后稿数变了 " + before + "→" + s.PresCount();
            }

            int mid = s.PresCount();
            if (!s.TryReadFullPage(out PptHtmlReadResult page, out err))
            {
                return "读包失败: " + err;
            }

            if (s.PresCount() != mid)
            {
                return "读包开了第二份稿 " + mid + "→" + s.PresCount();
            }

            List<PptHtmlShapeNode> charts = FindAll(page, "chart");
            return charts.Count < 1
                ? "读包没有 chart"
                : ContentAssert.ChartDataMismatch(charts[0], OoxCats, OoxVals, "pie");
        }

        private static string RunNegBadId(ContentGroupSession s)
        {
            HashSet<string> temp = ContentGroupSession.SnapshotEwTemp();
            OoxDeck deck;
            string err = SeedDeck(s, OoxPos.Mid, out deck);
            if (err != null)
            {
                return err;
            }

            int pages = s.SlideCount();
            bool ok = s.TryApply(
                ContentHtml.ChartUpdate("999001", OoxCats2, OoxVals2),
                false,
                out _,
                out err);
            if (ok)
            {
                return "坏 ShapeId 不该成功";
            }

            if (s.SlideCount() != pages)
            {
                return "失败路径改了页数";
            }

            if (!s.PresAlive())
            {
                return "失败路径打死原稿";
            }

            return ContentGroupSession.LeftoverEwTemp(temp) ?? AssertSiblings(s, deck);
        }

        private sealed class OoxDeck
        {
            public string CurId;
            public string A1Id;
            public string A2Id;
            public int StartCount;
        }

        private static string SeedDeck(ContentGroupSession s, OoxPos pos, out OoxDeck deck)
        {
            deck = new OoxDeck
            {
                CurId = s.SlideId,
                StartCount = s.SlideCount()
            };
            if (!s.TryApply(
                    ContentHtml.TextboxCreate("KEEP-CUR", left: 70, top: 4, width: 26, height: 8),
                    true,
                    out _,
                    out string err))
            {
                return "铺 KEEP-CUR 失败: " + err;
            }

            if (!s.TryAddBlankSlide(out deck.A1Id, out err)
                || !s.TrySelectSlide(deck.A1Id, out err)
                || !s.TryApply(
                    ContentHtml.TextboxCreate("KEEP-A1", left: 8, top: 8, width: 26, height: 8),
                    true,
                    out _,
                    out err))
            {
                return "铺 KEEP-A1 失败: " + err;
            }

            if (!s.TryAddBlankSlide(out deck.A2Id, out err)
                || !s.TrySelectSlide(deck.A2Id, out err)
                || !s.TryApply(
                    ContentHtml.TextboxCreate("KEEP-A2", left: 8, top: 8, width: 26, height: 8),
                    true,
                    out _,
                    out err))
            {
                return "铺 KEEP-A2 失败: " + err;
            }

            string target = deck.CurId;
            if (pos == OoxPos.Mid)
            {
                target = deck.A1Id;
            }
            else if (pos == OoxPos.Last)
            {
                target = deck.A2Id;
            }

            return s.TrySelectSlide(target, out err) ? null : err;
        }

        private static string AssertPipe(
            ContentGroupSession s,
            OoxDeck deck,
            HashSet<string> tempBefore,
            int startPres,
            bool requireSlim)
        {
            if (!s.PresAlive())
            {
                return "原稿死了";
            }

            if (s.SlideCount() != deck.StartCount + 2)
            {
                return "页数变了 期望 " + (deck.StartCount + 2) + " 实际 " + s.SlideCount();
            }

            if (startPres > 0 && s.PresCount() != startPres)
            {
                return "留下第二份稿 前=" + startPres + " 后=" + s.PresCount();
            }

            string leftover = ContentGroupSession.LeftoverEwTemp(tempBefore);
            if (leftover != null)
            {
                return leftover;
            }

            if (requireSlim && !HasSlimLog(s))
            {
                return "没看到瘦包日志" + DumpWarn(s);
            }

            return AssertSiblings(s, deck);
        }

        private static string AssertSiblings(ContentGroupSession s, OoxDeck deck)
        {
            string here = s.SlideId;
            string e = ExpectTextOn(s, deck.CurId, "KEEP-CUR")
                ?? ExpectTextOn(s, deck.A1Id, "KEEP-A1")
                ?? ExpectTextOn(s, deck.A2Id, "KEEP-A2");
            string back;
            if (!s.TrySelectSlide(here, out back) && e == null)
            {
                return back;
            }

            return e;
        }

        private static string ExpectTextOn(ContentGroupSession s, string slideId, string text)
        {
            if (!s.TrySelectSlide(slideId, out string err))
            {
                return "切到 " + text + " 页失败: " + err;
            }

            if (!s.TryReadPage(out PptHtmlReadResult page, out err))
            {
                return "读 " + text + " 页失败: " + err;
            }

            return ContentGroupSession.FindTextTop(page, text) == null
                ? "邻页丢了 " + text
                : null;
        }

        private static bool HasSlimLog(ContentGroupSession s)
        {
            if (s.LastWarnings == null)
            {
                return false;
            }

            for (int i = 0; i < s.LastWarnings.Count; i++)
            {
                string w = s.LastWarnings[i] ?? "";
                if (w.IndexOf("已瘦成一页+主题", StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string DumpWarn(ContentGroupSession s)
        {
            if (s.LastWarnings == null || s.LastWarnings.Count == 0)
            {
                return "";
            }

            return " warnings=[" + string.Join("; ", s.LastWarnings) + "]";
        }

        private static string TextHas(PptHtmlReadResult page, string text)
        {
            return ContentGroupSession.FindTextTop(page, text) == null ? "缺文案 " + text : null;
        }

        private static List<PptHtmlShapeNode> FindAll(PptHtmlReadResult result, string type)
        {
            var list = new List<PptHtmlShapeNode>();
            CollectType(result?.Shapes, type, list);
            return list;
        }

        private static void CollectType(IList<PptHtmlShapeNode> nodes, string type, List<PptHtmlShapeNode> into)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                PptHtmlShapeNode n = nodes[i];
                if (n == null)
                {
                    continue;
                }

                if (string.Equals(n.ShapeType, type, StringComparison.OrdinalIgnoreCase))
                {
                    into.Add(n);
                }

                CollectType(n.Children, type, into);
            }
        }

        private static string FindChartId(IList<string> ids)
        {
            if (ids == null)
            {
                return null;
            }

            for (int i = ids.Count - 1; i >= 0; i--)
            {
                if (!string.IsNullOrEmpty(ids[i]))
                {
                    return ids[i];
                }
            }

            return null;
        }

        private static void AddOox(List<ContentCase> list, ref int page, string name, Func<ContentGroupSession, string> run)
        {
            page++;
            list.Add(new ContentCase
            {
                Index = list.Count + 1,
                Batch = 1,
                Page = page,
                Name = name,
                CreateHtml = a => ContentHtml.Section(""),
                GroupRun = run
            });
        }
    }
}
