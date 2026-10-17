using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using PptChartRoundtripTest;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 把 PPT 结构 300 + 属性 250 改成文字 ChartConfig，在已打开的 Word / WPS 文字稿上建、换、抽。
    /// 换图种或系列数不一致时删掉再建。不 Quit 用户本来开着的稿。
    /// 过滤：--case=wh-ppt 全套；--case=wh-ppt-sink-line 只跑名字里含 sink-line 的条。
    /// </summary>
    internal static class WhPptSuite
    {
        public static int Target
        {
            get { return ChartCaseCatalog.Total + ChartAttrCatalog.Total; }
        }

        public static void Run(TestRun run, string host, string caseFilter)
        {
            string nameFilter = "";
            const string prefix = "wh-ppt";
            if (!string.IsNullOrEmpty(caseFilter)
                && caseFilter.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                && caseFilter.Length > prefix.Length)
            {
                nameFilter = caseFilter.Substring(prefix.Length).TrimStart('-');
            }

            var cases = new List<ChartCase>();
            cases.AddRange(ChartCaseCatalog.Build());
            cases.AddRange(ChartAttrCatalog.Build());
            if (!string.IsNullOrEmpty(nameFilter))
            {
                cases = cases.Where(c => c.Name != null && c.Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();
            }

            Console.WriteLine("文字图表套件 " + host + "  条数=" + cases.Count
                + (string.IsNullOrEmpty(nameFilter) ? "" : "  过滤=" + nameFilter));

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
                for (int i = 0; i < cases.Count; i++)
                {
                    RunOne(run, session, channel, cases[i]);
                    if ((i + 1) % 25 == 0)
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

        private static void RunOne(TestRun run, WdSession session, string channel, ChartCase one)
        {
            string tag = "wh-ppt-" + (one.Name ?? "?");
            string finalHtml = one.IsReplace ? one.ReplaceHtml : one.CreateHtml;
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
                    if (!WhPptChartXml.TryExport(one.CreateHtml, out string baseXml, out error))
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
