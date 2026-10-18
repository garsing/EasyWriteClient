using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using PptChartRoundtripTest;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 把 PPT 结构 300 + 属性 250 改成文字 ChartConfig，在已打开的 Word / WPS 文字稿上建、换、抽。
    /// 换图种或系列数不一致时删掉再建。不 Quit 用户本来开着的稿。
    /// 过滤：--case=wh-ppt 全套；--case=wh-ppt-sink-line 只跑名字里含 sink-line 的条。
    /// --wh-ids=158,167 只跑套件序号（1–550）。同名用例会重复出现，不能靠名字区分。
    /// </summary>
    internal static class WhPptSuite
    {
        public static int Target
        {
            get { return ChartCaseCatalog.Total + ChartAttrCatalog.Total; }
        }

        /// <param name="startAfter">这条的名字跑完后从下一条开始。空则从头。</param>
        /// <param name="onlyIds">套件序号，逗号分隔。空则不按序号收窄。</param>
        /// <summary>
        /// 同一条建图，只换 ChartStyle，打印请求值和读回值。不改 550 套件的期望。
        /// </summary>
        public static void StyleSweep(TestRun run, string host)
        {
            // WPS 文字只留 101–148。请求 1–48 时读回是原号加 100；这个区间外收成默认 102。
            int[] styles = { 2, 10, 48, 100, 101, 148, 149, 201, 286 };
            ChartCase combo = null;
            ChartCase column = null;
            foreach (ChartCase one in ChartAttrCatalog.Build())
            {
                if (one.IsReplace || one.Name == null)
                {
                    continue;
                }

                if (one.Name.EndsWith("combo-style-scale-color", StringComparison.OrdinalIgnoreCase))
                {
                    combo = one;
                }
                else if (one.Name.EndsWith("chart-style", StringComparison.OrdinalIgnoreCase))
                {
                    column = one;
                }
            }

            if (combo == null || column == null)
            {
                run.CaseFail("wh-ppt-style-sweep", "找不到样式探针底图");
                return;
            }

            var session = new WdSession(host);
            session.DismissOrphanExcelOnClose = !session.IsWps;
            string path = session.NewDocxPath();
            ToolResult opened = session.Open(path, true, session.App);
            if (opened == null || !opened.Success)
            {
                run.CaseFail("wh-ppt-style-sweep", opened == null ? "打开空白稿失败" : (opened.Error ?? "打开空白稿失败"));
                return;
            }

            string channel = WdSession.Str(opened, "channel_id");
            try
            {
                SweepOne(run, session, channel, "combo", combo.CreateHtml, styles);
                SweepOne(run, session, channel, "column", column.CreateHtml, new[] { 10, 48, 102, 201, 286 });
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static void SweepOne(TestRun run, WdSession session, string channel, string kind, string html, int[] styles)
        {
            for (int i = 0; i < styles.Length; i++)
            {
                int style = styles[i];
                string swapped = html.Replace("data-chart-style=\"286\"", "data-chart-style=\"" + style.ToString() + "\"");
                if (swapped.IndexOf("data-chart-style=\"" + style.ToString() + "\"", StringComparison.Ordinal) < 0)
                {
                    Console.WriteLine("  STYLE " + kind + " " + style + " -> 换号失败");
                    run.CaseFail("wh-ppt-style-" + kind + "-" + style, "HTML 里没有 data-chart-style");
                    continue;
                }

                if (!WhPptChartXml.TryExport(swapped, out string xml, out string error))
                {
                    Console.WriteLine("  STYLE " + kind + " " + style + " -> 转 XML 失败 " + error);
                    run.CaseFail("wh-ppt-style-" + kind + "-" + style, error ?? "转 XML 失败");
                    continue;
                }

                string chartId = Create(session, channel, xml, out error);
                if (chartId == null)
                {
                    Console.WriteLine("  STYLE " + kind + " " + style + " -> 建图失败 " + error);
                    run.CaseFail("wh-ppt-style-" + kind + "-" + style, error ?? "建图失败");
                    continue;
                }

                try
                {
                    ToolResult extracted = session.ExtractChartXml(ExtractArgs(channel, chartId));
                    string back = extracted == null ? null : WdSession.Str(extracted, "xml_content");
                    string got = "(无)";
                    if (!string.IsNullOrEmpty(back))
                    {
                        XElement el = XDocument.Parse(back).Descendants()
                            .FirstOrDefault(e => e.Name.LocalName == "ChartStyle");
                        got = el == null || string.IsNullOrEmpty(el.Value) ? "(null)" : el.Value.Trim();
                    }
                    else
                    {
                        got = extracted == null ? "抽图空" : (extracted.Error ?? "无 xml");
                    }

                    Console.WriteLine("  STYLE " + kind + " 请求 " + style + " 读回 " + got);
                    run.CaseOk("wh-ppt-style-" + kind + "-" + style);
                }
                finally
                {
                    session.DeleteChart(DelArgs(channel, chartId));
                }
            }
        }

        public static void Run(TestRun run, string host, string caseFilter, string startAfter, string onlyIds)
        {
            string nameFilter = "";
            const string prefix = "wh-ppt";
            if (!string.IsNullOrEmpty(caseFilter)
                && caseFilter.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                && caseFilter.Length > prefix.Length)
            {
                nameFilter = caseFilter.Substring(prefix.Length).TrimStart('-');
            }

            var built = new List<ChartCase>();
            built.AddRange(ChartCaseCatalog.Build());
            built.AddRange(ChartAttrCatalog.Build());
            var idSet = ParseIds(onlyIds);
            var cases = new List<ChartCase>();
            var suiteIds = new List<int>();
            for (int n = 0; n < built.Count; n++)
            {
                int suiteId = n + 1;
                if (idSet.Count > 0 && !idSet.Contains(suiteId))
                {
                    continue;
                }

                if (!string.IsNullOrEmpty(nameFilter)
                    && (built[n].Name == null || built[n].Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) < 0))
                {
                    continue;
                }

                cases.Add(built[n]);
                suiteIds.Add(suiteId);
            }

            int start = 0;
            if (!string.IsNullOrEmpty(startAfter))
            {
                int found = -1;
                for (int n = 0; n < cases.Count; n++)
                {
                    if (string.Equals(cases[n].Name, startAfter, StringComparison.OrdinalIgnoreCase))
                    {
                        found = n;
                        break;
                    }
                }

                if (found < 0)
                {
                    run.CaseFail("wh-ppt-from", "找不到起点 " + startAfter);
                    return;
                }

                start = found + 1;
            }

            run.TotalTarget = cases.Count - start;
            Console.WriteLine("文字图表套件 " + host + "  条数=" + (cases.Count - start)
                + (start == 0 ? "" : "  从过滤后第 " + (start + 1) + " 条起，上一条 " + startAfter)
                + (string.IsNullOrEmpty(nameFilter) ? "" : "  过滤=" + nameFilter)
                + (idSet.Count == 0 ? "" : "  编号=" + onlyIds));

            var session = new WdSession(host);
            session.DismissOrphanExcelOnClose = !session.IsWps;
            string path = session.NewDocxPath();
            ToolResult opened = session.Open(path, true, session.App);
            if (opened == null || !opened.Success)
            {
                run.CaseFail("wh-ppt-open", opened == null ? "打开空白稿失败" : (opened.Error ?? "打开空白稿失败"));
                return;
            }

            string channel = WdSession.Str(opened, "channel_id");
            try
            {
                for (int i = start; i < cases.Count; i++)
                {
                    Console.WriteLine("  #" + suiteIds[i] + " " + (cases[i].Name ?? "?"));
                    RunOneBounded(run, session, ref channel, ref path, cases[i]);
                    if ((i + 1) % 25 == 0 || i == cases.Count - 1)
                    {
                        Console.WriteLine("  … " + (i + 1) + "/" + cases.Count
                            + " 过=" + run.CasesPassed + " 败=" + run.CasesFailed);
                    }
                }
            }
            finally
            {
                session.CloseDocByPath(path);
            }
        }

        private static HashSet<int> ParseIds(string onlyIds)
        {
            var set = new HashSet<int>();
            if (string.IsNullOrWhiteSpace(onlyIds))
            {
                return set;
            }

            string[] parts = onlyIds.Split(new[] { ',', ' ', ';' }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < parts.Length; i++)
            {
                int id;
                if (int.TryParse(parts[i].Trim(), out id) && id > 0)
                {
                    set.Add(id);
                }
            }

            return set;
        }

        /// <summary>
        /// 单条超过 100 秒就关掉卡住的 Desktop 和无窗口 Excel，再开一份空白稿继续。
        /// 管道读取消停不住已经挂起的调用。
        /// </summary>
        private static void RunOneBounded(
            TestRun run,
            WdSession session,
            ref string channel,
            ref string path,
            ChartCase one)
        {
            string ch = channel;
            int mark = run.CasesPassed + run.CasesFailed;
            Exception boom = null;
            var done = new ManualResetEventSlim(false);
            var worker = new Thread(() =>
            {
                try
                {
                    RunOne(run, session, ch, one);
                }
                catch (Exception ex)
                {
                    boom = ex;
                }
                finally
                {
                    done.Set();
                }
            });
            worker.IsBackground = true;
            worker.Start();
            if (done.Wait(TimeSpan.FromSeconds(100)))
            {
                if (boom != null)
                {
                    run.CaseFail("wh-ppt-" + (one.Name ?? "?"), boom.Message);
                }

                return;
            }

            Console.WriteLine("  [超时] " + (one.Name ?? "?") + " 超过 100 秒，重启 Desktop 后继续");
            StopStuckHost();
            if (!done.Wait(TimeSpan.FromSeconds(20)))
            {
                if (run.CasesPassed + run.CasesFailed == mark)
                {
                    run.CaseFail("wh-ppt-" + (one.Name ?? "?"), "超时，调用未返回");
                }
            }
            else if (boom != null && run.CasesPassed + run.CasesFailed == mark)
            {
                run.CaseFail("wh-ppt-" + (one.Name ?? "?"), boom.Message);
            }

            Reopen(session, ref channel, ref path, run);
        }

        private static void StopStuckHost()
        {
            foreach (Process p in Process.GetProcessesByName("EasyWriteClient.Desktop"))
            {
                try
                {
                    p.Kill();
                }
                catch (Exception)
                {
                }
            }

            foreach (Process p in Process.GetProcessesByName("EXCEL"))
            {
                try
                {
                    if (p.MainWindowHandle == IntPtr.Zero)
                    {
                        p.Kill();
                    }
                }
                catch (Exception)
                {
                }
            }

            Thread.Sleep(1500);
        }

        private static void Reopen(WdSession session, ref string channel, ref string path, TestRun run)
        {
            string previous = path;
            if (!DesktopLauncher.EnsureAvailable())
            {
                run.CaseFail("wh-ppt-open", "超时后 Desktop 没有起来");
                return;
            }

            if (!string.IsNullOrEmpty(previous))
            {
                try
                {
                    session.CloseDocByPath(previous);
                }
                catch (Exception)
                {
                }
            }

            path = session.NewDocxPath();
            ToolResult opened = session.Open(path, true, session.App);
            if (opened == null || !opened.Success)
            {
                run.CaseFail("wh-ppt-open", opened == null ? "超时后重开空白稿失败" : (opened.Error ?? "超时后重开空白稿失败"));
                return;
            }

            channel = WdSession.Str(opened, "channel_id");
        }

        /// <summary>WPS 文字能原样留下的样式号。286 会收成默认 102，改请求和期望，Word / WPP 仍用 286。</summary>
        private const string WpsChartStyle = "148";

        private static void RunOne(TestRun run, WdSession session, string channel, ChartCase one)
        {
            string tag = "wh-ppt-" + (one.Name ?? "?");
            string createHtml = one.CreateHtml;
            string finalHtml = one.IsReplace ? one.ReplaceHtml : one.CreateHtml;
            if (session.IsWps)
            {
                ApplyWpsChartStyle(one, ref createHtml, ref finalHtml);
            }

            if (!WhPptChartXml.TryExport(finalHtml, out string finalXml, out string error))
            {
                run.CaseFail(tag, error ?? "HTML 转 ChartConfig 失败");
                return;
            }

            string chartId = null;
            try
            {
                if (one.IsReplace)
                {
                    if (!WhPptChartXml.TryExport(createHtml, out string baseXml, out error))
                    {
                        run.CaseFail(tag, error ?? "底图转 ChartConfig 失败");
                        return;
                    }

                    chartId = Create(session, channel, baseXml, out error);
                    if (chartId == null)
                    {
                        run.CaseFail(tag, "底图失败: " + error);
                        return;
                    }

                    if (SameShape(baseXml, finalXml))
                    {
                        ToolResult applied = session.ApplyChartFromXml(ApplyArgs(channel, chartId, finalXml));
                        if (applied == null || !applied.Success)
                        {
                            run.CaseFail(tag, "换数失败: " + (applied == null ? "空结果" : (applied.Error ?? "套用失败")));
                            return;
                        }
                    }
                    else
                    {
                        ToolResult deleted = session.DeleteChart(DelArgs(channel, chartId));
                        if (deleted == null || !deleted.Success)
                        {
                            run.CaseFail(tag, "删底图失败: " + (deleted == null ? "空结果" : (deleted.Error ?? "删除失败")));
                            return;
                        }

                        chartId = Create(session, channel, finalXml, out error);
                        if (chartId == null)
                        {
                            run.CaseFail(tag, "换结构再建失败: " + error);
                            return;
                        }
                    }
                }
                else
                {
                    chartId = Create(session, channel, finalXml, out error);
                    if (chartId == null)
                    {
                        run.CaseFail(tag, "建图失败: " + error);
                        return;
                    }
                }

                ToolResult extracted = session.ExtractChartXml(ExtractArgs(channel, chartId));
                string xml = extracted == null ? null : WdSession.Str(extracted, "xml_content");
                if (extracted == null || !extracted.Success || string.IsNullOrEmpty(xml))
                {
                    run.CaseFail(tag, "抽图失败: " + (extracted == null ? "空结果" : (extracted.Error ?? "无 xml_content")));
                    return;
                }

                string mismatch = WhPptChartXml.Match(one, xml);
                if (mismatch != null)
                {
                    run.CaseFail(tag, mismatch);
                    return;
                }

                run.CaseOk(tag);
            }
            finally
            {
                if (!string.IsNullOrEmpty(chartId))
                {
                    session.DeleteChart(DelArgs(channel, chartId));
                }
            }
        }

        private static void ApplyWpsChartStyle(ChartCase one, ref string createHtml, ref string finalHtml)
        {
            if (one?.Attrs == null)
            {
                return;
            }

            bool styleCase = false;
            for (int i = 0; i < one.Attrs.Count; i++)
            {
                AttrCheck check = one.Attrs[i];
                if (check == null
                    || !string.Equals(check.Key, "ChartStyle", StringComparison.Ordinal)
                    || !string.Equals(check.Expected, "286", StringComparison.Ordinal))
                {
                    continue;
                }

                check.Expected = WpsChartStyle;
                styleCase = true;
            }

            if (!styleCase)
            {
                return;
            }

            createHtml = SwapChartStyle(createHtml);
            finalHtml = SwapChartStyle(finalHtml);
        }

        private static string SwapChartStyle(string html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return html;
            }

            return html.Replace("data-chart-style=\"286\"", "data-chart-style=\"" + WpsChartStyle + "\"");
        }

        private static string Create(WdSession session, string channel, string xml, out string error)
        {
            error = null;
            string file = Path.Combine(Path.GetTempPath(), "ew-wh-ppt-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(file, xml);
            try
            {
                var args = new Dictionary<string, object>
                {
                    ["channel_id"] = channel,
                    ["path"] = file
                };
                ToolResult created = session.CreateChartFromXml(args);
                string chartId = created == null ? null : WdSession.Str(created, "chart_id");
                if (created == null || !created.Success || string.IsNullOrEmpty(chartId))
                {
                    error = created == null ? "空结果" : (created.Error ?? "建图失败");
                    return null;
                }

                return chartId;
            }
            finally
            {
                try
                {
                    File.Delete(file);
                }
                catch (Exception)
                {
                }
            }
        }

        private static bool SameShape(string leftXml, string rightXml)
        {
            try
            {
                XDocument left = XDocument.Parse(leftXml);
                XDocument right = XDocument.Parse(rightXml);
                return string.Equals(TypeOf(left), TypeOf(right), StringComparison.OrdinalIgnoreCase)
                    && ValueCount(left) == ValueCount(right);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static string TypeOf(XDocument doc)
        {
            XElement type = doc.Descendants().FirstOrDefault(e => e.Name.LocalName == "Type");
            return type == null ? "" : (type.Value ?? "").Trim();
        }

        private static int ValueCount(XDocument doc)
        {
            return doc.Descendants().Count(e =>
                e.Name.LocalName == "Column"
                && e.Attribute("type") != null
                && string.Equals(e.Attribute("type").Value, "value", StringComparison.OrdinalIgnoreCase));
        }

        private static Dictionary<string, object> ApplyArgs(string channel, string chartId, string xml)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["chart_id"] = chartId ?? "",
                ["scope"] = "all",
                ["xml_content"] = xml ?? ""
            };
        }

        private static Dictionary<string, object> DelArgs(string channel, string chartId)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["chart_id"] = chartId ?? ""
            };
        }

        private static Dictionary<string, object> ExtractArgs(string channel, string chartId)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channel,
                ["scope"] = "all",
                ["chart_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["chart_id"] = chartId ?? ""
                },
                ["path"] = Path.Combine(Path.GetTempPath(), "ew-wh-ppt-x-" + Guid.NewGuid().ToString("N") + ".xml")
            };
        }
    }
}
