using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Newtonsoft.Json.Linq;
using WordAddIn1.OpenFiles;
using WordAddIn1.PresentationHost;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptHtmlContentRoundtripTest
{
    /// <summary>打开易写7.24-2，按 XML 预搜金标测 query。</summary>
    internal static class FixtureSearchRunner
    {
        public static void Run(TestRun run, IList<string> nameFilters, bool useWpp)
        {
            string goldPath = FindFile("_query_gold.json");
            string pptx = FindFile("易写7.24-2.pptx");
            if (goldPath == null || pptx == null)
            {
                run.CaseFail(
                    "srch-fx-",
                    "找不到金标或真稿。gold=" + (goldPath ?? "无") + " pptx=" + (pptx ?? "无"));
                return;
            }

            List<FxCase> all = LoadGold(goldPath);
            List<FxCase> cases = Filter(all, nameFilters);
            Console.WriteLine((useWpp ? "真稿 query · WPP · " : "真稿 query · PPT · ") + pptx);
            Console.WriteLine("金标 " + all.Count + " 过滤后 " + cases.Count);
            if (cases.Count == 0)
            {
                return;
            }

            var bySlide = new Dictionary<int, List<FxCase>>();
            for (int i = 0; i < cases.Count; i++)
            {
                FxCase one = cases[i];
                if (!bySlide.TryGetValue(one.SlideId, out List<FxCase> list))
                {
                    list = new List<FxCase>();
                    bySlide[one.SlideId] = list;
                }

                list.Add(one);
            }

            if (useWpp)
            {
                RunWpp(run, pptx, bySlide);
            }
            else
            {
                RunPpt(run, pptx, bySlide);
            }
        }

        private static void RunPpt(
            TestRun run,
            string pptx,
            Dictionary<int, List<FxCase>> bySlide)
        {
            PowerPoint.Application app = null;
            PowerPoint.Presentation pres = null;
            try
            {
                app = new PowerPoint.Application();
                app.Visible = Office.MsoTriState.msoTrue;
                try
                {
                    app.DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone;
                }
                catch
                {
                }

                pres = app.Presentations.Open(
                    pptx,
                    Office.MsoTriState.msoTrue,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue);
                Console.WriteLine("已打开 页=" + pres.Slides.Count);
                bool ReadPpt(string slideId, out PptHtmlReadResult tree, out string readErr)
                {
                    return PptHtmlPowerPointReader.TryRead(
                        pres,
                        slideId,
                        "fx",
                        "ppt",
                        null,
                        false,
                        true,
                        out tree,
                        out readErr);
                }

                RunBySlide(run, bySlide, ReadPpt);
            }
            catch (Exception ex)
            {
                run.CaseFail("srch-fx-", "打开或搜索失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        pres.Saved = Office.MsoTriState.msoTrue;
                        pres.Close();
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null && app.Presentations.Count == 0)
                    {
                        app.Quit();
                    }
                }
                catch
                {
                }
            }
        }

        private static void RunWpp(
            TestRun run,
            string pptx,
            Dictionary<int, List<FxCase>> bySlide)
        {
            object app = null;
            object pres = null;
            try
            {
                if (!WppCom.TryCreateApplication(out app, out string startErr))
                {
                    run.CaseFail("srch-fx-", startErr);
                    return;
                }

                object presentations = WppCom.GetProperty(app, "Presentations");
                try
                {
                    pres = WppCom.Invoke(presentations, "Open", pptx, true, false, true);
                }
                catch
                {
                    pres = WppCom.Invoke(presentations, "Open", pptx);
                }

                if (pres == null)
                {
                    run.CaseFail("srch-fx-", "WPP 打开真稿失败");
                    return;
                }

                object slides = WppCom.GetProperty(pres, "Slides");
                int pageCount = Convert.ToInt32(WppCom.GetProperty(slides, "Count") ?? 0);
                Console.WriteLine("已打开 页=" + pageCount.ToString(CultureInfo.InvariantCulture));
                bool ReadWpp(string slideId, out PptHtmlReadResult tree, out string readErr)
                {
                    return PptHtmlWppReader.TryRead(
                        pres,
                        slideId,
                        "fx-wpp",
                        "wpp",
                        null,
                        false,
                        true,
                        out tree,
                        out readErr);
                }

                RunBySlide(run, bySlide, ReadWpp);
            }
            catch (Exception ex)
            {
                run.CaseFail("srch-fx-", "打开或搜索失败: " + ex.Message);
            }
            finally
            {
                try
                {
                    if (pres != null)
                    {
                        WppCom.TrySetProperty(pres, "Saved", true);
                        WppCom.Invoke(pres, "Close");
                    }
                }
                catch
                {
                }

                try
                {
                    if (app != null)
                    {
                        object presentations = WppCom.GetProperty(app, "Presentations");
                        int n = Convert.ToInt32(WppCom.GetProperty(presentations, "Count") ?? 0);
                        if (n == 0)
                        {
                            WppCom.Invoke(app, "Quit");
                        }
                    }
                }
                catch
                {
                }
            }
        }

        private delegate bool TryReadSlide(
            string slideId,
            out PptHtmlReadResult tree,
            out string readErr);

        private static void RunBySlide(
            TestRun run,
            Dictionary<int, List<FxCase>> bySlide,
            TryReadSlide read)
        {
            foreach (KeyValuePair<int, List<FxCase>> kv in bySlide)
            {
                string slideId = kv.Key.ToString(CultureInfo.InvariantCulture);
                Console.WriteLine("--- slideId=" + slideId + " " + kv.Value.Count + " 条 ---");
                if (!read(slideId, out PptHtmlReadResult tree, out string readErr) || tree == null)
                {
                    for (int i = 0; i < kv.Value.Count; i++)
                    {
                        run.CaseFail(kv.Value[i].Id, "读页失败: " + (readErr ?? "?"));
                    }

                    continue;
                }

                for (int i = 0; i < kv.Value.Count; i++)
                {
                    RunOne(run, tree.Shapes, kv.Value[i]);
                }
            }
        }

        public static bool WantsFixture(IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return false;
            }

            for (int i = 0; i < nameFilters.Count; i++)
            {
                string f = nameFilters[i] ?? "";
                if (f.IndexOf("srch-fx", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void RunOne(TestRun run, IList<PptHtmlShapeNode> roots, FxCase one)
        {
            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);
            var query = new List<object>();
            for (int i = 0; i < one.Query.Count; i++)
            {
                query.Add(new Dictionary<string, object>
                {
                    ["attr"] = one.Query[i].Attr,
                    ["pattern"] = one.Query[i].Pattern
                });
            }

            args["query"] = query;
            string match = one.Match ?? "";
            if (string.Equals(match, "all", StringComparison.OrdinalIgnoreCase))
            {
                match = "and";
            }
            else if (string.Equals(match, "any", StringComparison.OrdinalIgnoreCase))
            {
                match = "or";
            }

            if (one.Query.Count >= 2 || !string.IsNullOrEmpty(match))
            {
                args["match"] = match;
            }
            if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest request, out string parseErr))
            {
                run.CaseFail(one.Id, "解析 query 失败: " + parseErr);
                return;
            }

            if (!PptHtmlNodeSearch.TryMatch(
                    roots,
                    request,
                    out List<PptHtmlShapeNode> forest,
                    out int leafCount,
                    out int matchTotal,
                    out bool _,
                    out string matchErr))
            {
                run.CaseFail(
                    one.Id,
                    "TryMatch 失败: " + (matchErr ?? "?")
                    + " leaf=" + leafCount
                    + " total=" + matchTotal
                    + " 期望 " + one.MatchCount);
                return;
            }

            var actual = new HashSet<string>(StringComparer.Ordinal);
            CollectLeafIds(forest, actual);
            var expect = new HashSet<string>(one.ShapeIds ?? new List<string>(), StringComparer.Ordinal);
            string kind = CaseKind(one.Id);
            if (kind == "pic")
            {
                if (actual.Count > 0)
                {
                    run.CaseOk(one.Id + " COM picture n=" + actual.Count + " XML pic=" + expect.Count);
                    return;
                }

                run.CaseFail(one.Id, "COM 一张 picture 都没有");
                return;
            }

            if (kind == "any")
            {
                HashSet<string> textHits = MatchTextOnly(roots, one);
                List<string> missingText = MissingFrom(textHits, actual);
                if (missingText.Count == 0 && textHits.Count > 0)
                {
                    run.CaseOk(one.Id + " 正文OR⊆COM n=" + actual.Count);
                    return;
                }

                if (textHits.Count == 0 && actual.Count > 0)
                {
                    run.CaseOk(one.Id + " COM 有命中 n=" + actual.Count + "（XML 正文词 COM 未读到）");
                    return;
                }

                run.CaseFail(one.Id, "OR 未覆盖 COM 正文命中");
                return;
            }

            if (kind == "tb" || kind == "and")
            {
                List<string> extra = MissingFrom(actual, expect);
                if (extra.Count == 0)
                {
                    run.CaseOk(one.Id + " COM⊆XML n=" + actual.Count + "/" + expect.Count);
                    return;
                }

                run.CaseFail(
                    one.Id,
                    "COM 命中应是 XML 金标子集。多出 " + extra[0]
                    + " COM=" + actual.Count + " XML=" + expect.Count);
                return;
            }

            if (actual.SetEquals(expect) && leafCount == one.MatchCount)
            {
                run.CaseOk(one.Id + " n=" + leafCount);
                return;
            }

            run.CaseFail(
                one.Id,
                "期望 n=" + one.MatchCount + " ids=[" + JoinSorted(expect) + "]"
                + " 实际 n=" + leafCount + " ids=[" + JoinSorted(actual) + "]");
        }

        private static string CaseKind(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return "";
            }

            if (id.EndsWith("-type-pic", StringComparison.OrdinalIgnoreCase))
            {
                return "pic";
            }

            if (id.EndsWith("-type-tb", StringComparison.OrdinalIgnoreCase))
            {
                return "tb";
            }

            if (id.EndsWith("-text-any", StringComparison.OrdinalIgnoreCase))
            {
                return "exact";
            }

            if (id.EndsWith("-any", StringComparison.OrdinalIgnoreCase))
            {
                return "any";
            }

            if (id.EndsWith("-and", StringComparison.OrdinalIgnoreCase))
            {
                return "and";
            }

            return "exact";
        }

        private static HashSet<string> MatchTextOnly(IList<PptHtmlShapeNode> roots, FxCase one)
        {
            var ids = new HashSet<string>(StringComparer.Ordinal);
            FxClause text = null;
            if (one.Query != null)
            {
                for (int i = 0; i < one.Query.Count; i++)
                {
                    if (string.Equals(one.Query[i].Attr, "text", StringComparison.OrdinalIgnoreCase))
                    {
                        text = one.Query[i];
                        break;
                    }
                }
            }

            if (text == null)
            {
                return ids;
            }

            var args = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase)
            {
                ["query"] = new List<object>
                {
                    new Dictionary<string, object>
                    {
                        ["attr"] = "text",
                        ["pattern"] = text.Pattern
                    }
                }
            };
            if (!PptHtmlNodeSearch.TryParseRequest(args, out PptHtmlSearchRequest request, out _))
            {
                return ids;
            }

            if (!PptHtmlNodeSearch.TryMatch(roots, request, out List<PptHtmlShapeNode> forest, out _, out _, out _, out _))
            {
                return ids;
            }

            CollectLeafIds(forest, ids);
            return ids;
        }

        private static List<string> MissingFrom(HashSet<string> need, HashSet<string> have)
        {
            var missing = new List<string>();
            foreach (string id in need)
            {
                if (!have.Contains(id))
                {
                    missing.Add(id);
                }
            }

            return missing;
        }

        private static void CollectLeafIds(IList<PptHtmlShapeNode> nodes, HashSet<string> ids)
        {
            if (nodes == null)
            {
                return;
            }

            for (int i = 0; i < nodes.Count; i++)
            {
                CollectLeafIds(nodes[i], ids);
            }
        }

        private static void CollectLeafIds(PptHtmlShapeNode node, HashSet<string> ids)
        {
            if (node == null)
            {
                return;
            }

            bool group = string.Equals(node.ShapeType, "group", StringComparison.OrdinalIgnoreCase);
            if (group)
            {
                CollectLeafIds(node.Children, ids);
                return;
            }

            if (!string.IsNullOrEmpty(node.ShapeId))
            {
                ids.Add(node.ShapeId);
            }
        }

        private static string JoinSorted(HashSet<string> set)
        {
            var list = new List<string>(set);
            list.Sort(StringComparer.Ordinal);
            if (list.Count > 12)
            {
                return string.Join(",", list.GetRange(0, 12).ToArray()) + ",…共" + list.Count;
            }

            return string.Join(",", list.ToArray());
        }

        private static List<FxCase> LoadGold(string path)
        {
            var list = new List<FxCase>();
            JArray arr = JArray.Parse(File.ReadAllText(path));
            foreach (JToken tok in arr)
            {
                JObject o = tok as JObject;
                if (o == null)
                {
                    continue;
                }

                var one = new FxCase
                {
                    Id = (string)o["id"],
                    Page = (int)o["page"],
                    SlideId = (int)o["slideId"],
                    Title = (string)o["title"],
                    Match = (string)o["match"] ?? "all",
                    MatchCount = (int)o["match_count"],
                    Query = new List<FxClause>(),
                    ShapeIds = new List<string>()
                };
                JArray q = o["query"] as JArray;
                if (q != null)
                {
                    foreach (JToken c in q)
                    {
                        JObject cq = c as JObject;
                        if (cq == null)
                        {
                            continue;
                        }

                        one.Query.Add(new FxClause
                        {
                            Attr = (string)cq["attr"],
                            Pattern = (string)cq["pattern"]
                        });
                    }
                }

                JArray ids = o["shape_ids"] as JArray;
                if (ids != null)
                {
                    foreach (JToken id in ids)
                    {
                        string s = (string)id;
                        if (!string.IsNullOrEmpty(s))
                        {
                            one.ShapeIds.Add(s);
                        }
                    }
                }

                list.Add(one);
            }

            return list;
        }

        private static List<FxCase> Filter(List<FxCase> all, IList<string> nameFilters)
        {
            if (nameFilters == null || nameFilters.Count == 0)
            {
                return all;
            }

            bool onlyPrefix = true;
            for (int i = 0; i < nameFilters.Count; i++)
            {
                if (!string.Equals(nameFilters[i], "srch-fx-", StringComparison.OrdinalIgnoreCase)
                    && !string.Equals(nameFilters[i], "srch-fx", StringComparison.OrdinalIgnoreCase))
                {
                    onlyPrefix = false;
                    break;
                }
            }

            if (onlyPrefix)
            {
                return all;
            }

            var list = new List<FxCase>();
            for (int i = 0; i < all.Count; i++)
            {
                FxCase one = all[i];
                for (int f = 0; f < nameFilters.Count; f++)
                {
                    if (one.Id != null
                        && one.Id.IndexOf(nameFilters[f], StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        list.Add(one);
                        break;
                    }
                }
            }

            return list;
        }

        private static string FindFile(string name)
        {
            string dir = AppDomain.CurrentDomain.BaseDirectory;
            for (int i = 0; i < 10 && !string.IsNullOrEmpty(dir); i++)
            {
                string[] hits =
                {
                    Path.Combine(dir, name),
                    Path.Combine(dir, "docs", "测试", "ppt测试", "ppt搜索", "页内节点搜索", name),
                    Path.Combine(dir, "docs", "测试", "ppt搜索", "页内节点搜索", name)
                };
                for (int h = 0; h < hits.Length; h++)
                {
                    if (File.Exists(hits[h]))
                    {
                        return Path.GetFullPath(hits[h]);
                    }
                }

                dir = Directory.GetParent(dir)?.FullName;
            }

            return null;
        }

        private sealed class FxCase
        {
            public string Id;
            public int Page;
            public int SlideId;
            public string Title;
            public string Match;
            public int MatchCount;
            public List<FxClause> Query;
            public List<string> ShapeIds;
        }

        private sealed class FxClause
        {
            public string Attr;
            public string Pattern;
        }
    }
}
