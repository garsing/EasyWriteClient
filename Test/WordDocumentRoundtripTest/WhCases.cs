using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using WordAddIn1;

namespace WordDocumentRoundtripTest
{
    /// <summary>
    /// 第 7 波：ChartConfig 建图、抽 XML、套格式/数据、删图重建。
    /// Word 正例须成功。WPS 四件明确 unsupported，同一套 ID，不 SKIP。
    /// </summary>
    internal static class WhCases
    {
        public const int Target = 44;

        private const string AfterSent = "易写图前句。";
        private const string TitleA = "易写图销量";
        private const string TitleB = "易写图改过";
        private const string TitleLine = "易写图折线";
        private const string TitlePie = "易写图饼";

        private static readonly string[] HostChartIds =
        {
            "wh-create-ok", "wh-create-id", "wh-create-fields", "wh-create-text",
            "wh-create-filename", "wh-create-line", "wh-create-pie", "wh-create-two",
            "wh-create-keep-sent",
            "wh-xext-ok", "wh-xext-file", "wh-xext-data", "wh-xext-format",
            "wh-xext-no-sel", "wh-xext-unk", "wh-xext-ctx", "wh-xext-bad-scope",
            "wh-apply-fmt-ok", "wh-apply-data-ok", "wh-apply-data-mismatch",
            "wh-apply-unk", "wh-apply-echo", "wh-apply-keep-sent",
            "wh-del-ok", "wh-del-gone", "wh-del-keep-sent", "wh-del-twice",
            "wh-del-unk", "wh-rebuild-ok"
        };

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
                RunCreateAndMutate(run, s);
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

            s.ActivateDoc(path);
            s.GetContent(WdSession.Str(opened, "channel_id"));
            return path;
        }

        private static void RunNegEarly(TestRun run, WdSession s)
        {
            string xml = WriteXml(ColumnXml(TitleA));
            ToolResult noCh = s.CreateChartFromXml(CreateArgs(null, xml));
            if (!noCh.Success
                && (Contains(noCh.Error, "打开")
                    || Contains(noCh.Error, "unsupported")
                    || Contains(noCh.Error, "browser")
                    || Contains(noCh.Error, "渠道")))
            {
                run.CaseOk("wh-neg-no-channel");
            }
            else if (noCh.Success)
            {
                run.CaseSkip("wh-neg-no-channel", "现场已有默认文字渠，不关用户稿");
            }
            else
            {
                run.CaseFail("wh-neg-no-channel", noCh.Error ?? "应失败或跳过");
            }

            ExpectFail(run, "wh-neg-unknown-id", s.CreateChartFromXml(CreateArgs("w999", xml)), "未知");
            ExpectFail(run, "wh-neg-create-no-path",
                s.CreateChartFromXml(new Dictionary<string, object>()),
                "path");
            ExpectFail(run, "wh-neg-create-missing",
                s.CreateChartFromXml(CreateArgs(null, Path.Combine(Path.GetTempPath(), "ew-wh-missing-" + Guid.NewGuid().ToString("N") + ".xml"))),
                "不存在");
            ExpectFail(run, "wh-neg-create-bad-xml",
                s.CreateChartFromXml(CreateArgs(null, WriteXml("<NotChart><Chart><Type>column_clustered</Type></Chart></NotChart>"))),
                "ChartConfig");
            ExpectFail(run, "wh-neg-create-no-type",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Title>x</Title><DataTable><Header><Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header><Rows><Row>2023,10</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Type");
            ExpectFail(run, "wh-neg-create-bad-type",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(TypedXml("bubble", TitleA)))),
                "无效");
            ExpectFail(run, "wh-neg-del-no-id",
                s.DeleteChart(new Dictionary<string, object>()),
                "编号");
            ExpectFail(run, "wh-neg-apply-no-id",
                s.ApplyChartFromXml(new Dictionary<string, object>
                {
                    ["xml_content"] = FormatXml(TitleB)
                }),
                "编号");
            ExpectFail(run, "wh-neg-apply-no-xml",
                s.ApplyChartFromXml(new Dictionary<string, object> { ["chart_id"] = "C_00001" }),
                "path");
            ExpectFail(run, "wh-neg-apply-bad-scope",
                s.ApplyChartFromXml(new Dictionary<string, object>
                {
                    ["chart_id"] = "C_00001",
                    ["scope"] = "foo",
                    ["xml_content"] = FormatXml(TitleB)
                }),
                "scope");
        }

        private static void RunCreateAndMutate(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, HostChartIds))
            {
                return;
            }

            if (!s.TypePlain(path, AfterSent))
            {
                FailMany(run, "COM 写图前句失败", HostChartIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);

            if (s.IsWps)
            {
                RunWpsUnsupported(run, s, id);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult created = s.CreateChartFromXml(CreateArgs(id, WriteXml(ColumnXml(TitleA))));
            string chartId = WdSession.Str(created, "chart_id");
            Ok(run, "wh-create-ok", created.Success, created.Error ?? "建图失败");
            Ok(run, "wh-create-id",
                created.Success && chartId.StartsWith("C_", StringComparison.Ordinal),
                created.Success ? "无 C_" : (created.Error ?? ""));
            Ok(run, "wh-create-fields",
                created.Success
                    && WdSession.Bool(created, "chart_created")
                    && Contains(WdSession.Str(created, "chart_type"), "column"),
                created.Success ? "回包缺 chart_created / type" : (created.Error ?? ""));
            Ok(run, "wh-create-text",
                created.Success && HasChart(s, id),
                created.Success ? "正文无 <Chart C_>" : (created.Error ?? ""));
            Ok(run, "wh-create-keep-sent",
                created.Success && StillHas(s, id, AfterSent),
                created.Success ? "建图丢了前句" : (created.Error ?? ""));
            if (!created.Success || string.IsNullOrEmpty(chartId))
            {
                FailMany(run, created.Error ?? "建图无 chart_id",
                    "wh-create-filename", "wh-create-line", "wh-create-pie", "wh-create-two",
                    "wh-xext-ok", "wh-xext-file", "wh-xext-data", "wh-xext-format",
                    "wh-xext-no-sel", "wh-xext-unk", "wh-xext-ctx", "wh-xext-bad-scope",
                    "wh-apply-fmt-ok", "wh-apply-data-ok", "wh-apply-data-mismatch",
                    "wh-apply-unk", "wh-apply-echo", "wh-apply-keep-sent",
                    "wh-del-ok", "wh-del-gone", "wh-del-keep-sent", "wh-del-twice",
                    "wh-del-unk", "wh-rebuild-ok");
                s.CloseDocByPath(path);
                return;
            }

            string extractPath = Path.Combine(Path.GetTempPath(), "ew-wh-x-" + Guid.NewGuid().ToString("N") + ".xml");
            ToolResult xall = s.ExtractChartXml(ExtractArgs(id, chartId, extractPath, "all"));
            Ok(run, "wh-xext-ok", xall.Success, xall.Error ?? "抽图失败");
            Ok(run, "wh-xext-file",
                xall.Success && (File.Exists(extractPath) || !string.IsNullOrEmpty(WdSession.Str(xall, "xml_content"))),
                xall.Success ? "抽图无文件也无 xml_content" : (xall.Error ?? ""));

            ToolResult xdata = s.ExtractChartXml(ExtractArgs(id, chartId, null, "data"));
            string dataXml = WdSession.Str(xdata, "xml_content");
            Ok(run, "wh-xext-data",
                xdata.Success && (Contains(dataXml, "Rows") || Contains(dataXml, "Header")),
                xdata.Success ? "scope=data 无 Header/Rows" : (xdata.Error ?? ""));

            ToolResult xfmt = s.ExtractChartXml(ExtractArgs(id, chartId, null, "format"));
            Ok(run, "wh-xext-format",
                xfmt.Success && Contains(WdSession.Str(xfmt, "xml_content"), "ChartConfig"),
                xfmt.Success ? "scope=format 无 ChartConfig" : (xfmt.Error ?? ""));

            ExpectFail(run, "wh-xext-no-sel",
                s.ExtractChartXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wh-ns-" + Guid.NewGuid().ToString("N") + ".xml")
                }),
                "chart_selection");
            ExpectFail(run, "wh-xext-unk",
                s.ExtractChartXml(ExtractArgs(id, "C_99999", null, "all")),
                "找不到");
            ExpectFail(run, "wh-xext-ctx",
                s.ExtractChartXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["chart_selection"] = new Dictionary<string, object>
                    {
                        ["type"] = "by_context",
                        ["chart_id"] = chartId
                    }
                }),
                "by_context");
            ExpectFail(run, "wh-xext-bad-scope",
                s.ExtractChartXml(ExtractArgs(id, chartId, null, "foo")),
                "scope");

            ExpectFail(run, "wh-apply-data-mismatch",
                s.ApplyChartFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["chart_id"] = chartId,
                    ["scope"] = "data",
                    ["xml_content"] = MismatchXml()
                }),
                "chart_data_structure_mismatch");

            ToolResult applyFmt = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = chartId,
                ["scope"] = "format",
                ["xml_content"] = FormatXml(TitleB)
            });
            Ok(run, "wh-apply-fmt-ok", applyFmt.Success, applyFmt.Error ?? "套格式失败");
            Ok(run, "wh-apply-echo",
                applyFmt.Success
                    && WdSession.Bool(applyFmt, "chart_applied")
                    && Contains(WdSession.Str(applyFmt, "chart_id"), chartId),
                applyFmt.Success ? "回包缺 chart_applied / chart_id" : (applyFmt.Error ?? ""));
            Ok(run, "wh-apply-keep-sent",
                applyFmt.Success && StillHas(s, id, AfterSent),
                applyFmt.Success ? "套格式丢了前句" : (applyFmt.Error ?? ""));

            ToolResult applyData = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = chartId,
                ["scope"] = "data",
                ["xml_content"] = DataXml("11", "22", "33")
            });
            Ok(run, "wh-apply-data-ok", applyData.Success, applyData.Error ?? "套数据失败");

            ExpectFail(run, "wh-apply-unk",
                s.ApplyChartFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["chart_id"] = "C_99999",
                    ["xml_content"] = FormatXml(TitleB)
                }),
                "找不到");

            ToolResult line = s.CreateChartFromXml(CreateArgsFilename(id, WriteXml(TypedXml("line", TitleLine))));
            string lineId = WdSession.Str(line, "chart_id");
            Ok(run, "wh-create-filename",
                line.Success && StillHas(s, id, AfterSent),
                line.Success ? "filename 建图丢了前句" : (line.Error ?? ""));
            Ok(run, "wh-create-line",
                line.Success && Contains(WdSession.Str(line, "chart_type"), "line"),
                line.Success ? "折线 type 不对" : (line.Error ?? ""));
            Ok(run, "wh-create-two",
                line.Success && lineId.StartsWith("C_", StringComparison.Ordinal) && lineId != chartId && ChartCount(s, id) >= 2,
                line.Success ? "第二张 C_ 不对或正文只有一张" : (line.Error ?? ""));

            ToolResult pie = s.CreateChartFromXml(CreateArgs(id, WriteXml(TypedXml("pie3d", TitlePie))));
            Ok(run, "wh-create-pie",
                pie.Success && Contains(WdSession.Str(pie, "chart_type"), "pie"),
                pie.Success ? "饼图 type 不对" : (pie.Error ?? ""));

            ExpectFail(run, "wh-del-unk",
                s.DeleteChart(new Dictionary<string, object> { ["channel_id"] = id, ["chart_id"] = "C_99999" }),
                "找不到");

            ToolResult del = s.DeleteChart(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["table_id"] = chartId,
                ["chart_id"] = chartId
            });
            Ok(run, "wh-del-ok", del.Success, del.Error ?? "删图失败");
            Ok(run, "wh-del-gone",
                del.Success && !HasChartId(s, id, chartId),
                del.Success ? "删后正文还挂着该 C_" : (del.Error ?? ""));
            Ok(run, "wh-del-keep-sent",
                del.Success && StillHas(s, id, AfterSent) && (string.IsNullOrEmpty(lineId) || HasChartId(s, id, lineId)),
                del.Success ? "删第一张带走了前句或第二张" : (del.Error ?? ""));
            ExpectFail(run, "wh-del-twice",
                s.DeleteChart(new Dictionary<string, object> { ["channel_id"] = id, ["chart_id"] = chartId }),
                "找不到");

            ToolResult rebuild = s.CreateChartFromXml(CreateArgs(id, WriteXml(ColumnXml(TitleA))));
            Ok(run, "wh-rebuild-ok",
                rebuild.Success
                    && WdSession.Str(rebuild, "chart_id").StartsWith("C_", StringComparison.Ordinal)
                    && StillHas(s, id, AfterSent),
                rebuild.Success ? "重建无 C_ 或丢了前句" : (rebuild.Error ?? ""));

            s.CloseDocByPath(path);
        }

        private static void RunWpsUnsupported(TestRun run, WdSession s, string id)
        {
            ToolResult create = s.CreateChartFromXml(CreateArgs(id, WriteXml(ColumnXml(TitleA))));
            ExpectFail(run, "wh-create-ok", create, "unsupported");
            ExpectFail(run, "wh-create-id", create, "unsupported");
            ExpectFail(run, "wh-create-fields", create, "unsupported");
            ExpectFail(run, "wh-create-text", create, "unsupported");
            ExpectFail(run, "wh-create-keep-sent", create, "unsupported");
            ExpectFail(run, "wh-create-filename",
                s.CreateChartFromXml(CreateArgsFilename(id, WriteXml(TypedXml("line", TitleLine)))),
                "unsupported");
            ExpectFail(run, "wh-create-line", create, "unsupported");
            ExpectFail(run, "wh-create-pie", create, "unsupported");
            ExpectFail(run, "wh-create-two", create, "unsupported");

            var extract = s.ExtractChartXml(ExtractArgs(id, "C_00001", null, "all"));
            ExpectFail(run, "wh-xext-ok", extract, "unsupported");
            ExpectFail(run, "wh-xext-file", extract, "unsupported");
            ExpectFail(run, "wh-xext-data", extract, "unsupported");
            ExpectFail(run, "wh-xext-format", extract, "unsupported");
            ExpectFail(run, "wh-xext-no-sel", extract, "unsupported");
            ExpectFail(run, "wh-xext-unk", extract, "unsupported");
            ExpectFail(run, "wh-xext-ctx", extract, "unsupported");
            ExpectFail(run, "wh-xext-bad-scope", extract, "unsupported");

            var apply = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = "C_00001",
                ["xml_content"] = FormatXml(TitleB)
            });
            ExpectFail(run, "wh-apply-fmt-ok", apply, "unsupported");
            ExpectFail(run, "wh-apply-data-ok", apply, "unsupported");
            ExpectFail(run, "wh-apply-data-mismatch", apply, "unsupported");
            ExpectFail(run, "wh-apply-unk", apply, "unsupported");
            ExpectFail(run, "wh-apply-echo", apply, "unsupported");
            ExpectFail(run, "wh-apply-keep-sent", apply, "unsupported");

            var del = s.DeleteChart(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = "C_00001"
            });
            ExpectFail(run, "wh-del-ok", del, "unsupported");
            ExpectFail(run, "wh-del-gone", del, "unsupported");
            ExpectFail(run, "wh-del-keep-sent", del, "unsupported");
            ExpectFail(run, "wh-del-twice", del, "unsupported");
            ExpectFail(run, "wh-del-unk", del, "unsupported");
            ExpectFail(run, "wh-rebuild-ok", create, "unsupported");
        }

        private static void RunNegLate(TestRun run, WdSession s)
        {
            string xml = WriteXml(ColumnXml(TitleA));
            string xlsx = Path.Combine(Path.GetTempPath(), "ew-wh-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ToolResult xl = s.Open(xlsx, true, "excel");
            if (!xl.Success)
            {
                run.CaseFail("wh-neg-wrong-kind-excel", xl.Error ?? "建 xlsx 失败");
            }
            else
            {
                ExpectFail(run, "wh-neg-wrong-kind-excel",
                    s.CreateChartFromXml(CreateArgs(WdSession.Str(xl, "channel_id"), xml)), "excel");
                WdSession.CloseExcelByPath(WdSession.Str(xl, "path"));
            }

            string pptx = Path.Combine(Path.GetTempPath(), "ew-wh-" + Guid.NewGuid().ToString("N") + ".pptx");
            ToolResult ppt = s.Open(pptx, true, "powerpoint");
            if (!ppt.Success)
            {
                run.CaseFail("wh-neg-wrong-kind-ppt", ppt.Error ?? "建 pptx 失败");
            }
            else
            {
                ExpectFail(run, "wh-neg-wrong-kind-ppt",
                    s.CreateChartFromXml(CreateArgs(WdSession.Str(ppt, "channel_id"), xml)), "ppt");
                s.CloseDocByPath(WdSession.Str(ppt, "path"));
            }

            string closePath = s.NewDocxPath();
            ToolResult opened = s.Open(closePath, true, s.App);
            if (!opened.Success)
            {
                run.CaseFail("wh-neg-closed", opened.Error ?? "再建稿失败");
                return;
            }

            string cid = WdSession.Str(opened, "channel_id");
            if (!s.CloseDocByPath(closePath))
            {
                run.CaseFail("wh-neg-closed", "COM 关稿失败");
                return;
            }

            ToolResult closed = s.CreateChartFromXml(CreateArgs(cid, xml));
            if (closed != null && !closed.Success
                && (Contains(closed.Error, "关闭")
                    || Contains(closed.Error, "找不到")
                    || Contains(closed.Error, "未知")
                    || Contains(closed.Error, "unsupported")))
            {
                run.CaseOk("wh-neg-closed");
            }
            else if (closed != null && !closed.Success)
            {
                run.CaseFail("wh-neg-closed", closed.Error ?? "");
            }
            else
            {
                run.CaseFail("wh-neg-closed", "关稿后仍成功");
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
            Ok(run, "wh-no-quit", wordOk && wpsOk, wordDetail + " " + wpsDetail);
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

        private static Dictionary<string, object> ExtractArgs(string channelId, string chartId, string path, string scope)
        {
            var args = new Dictionary<string, object>
            {
                ["scope"] = scope ?? "all",
                ["chart_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["chart_id"] = chartId ?? ""
                }
            };
            if (channelId != null)
            {
                args["channel_id"] = channelId;
            }

            if (!string.IsNullOrEmpty(path))
            {
                args["path"] = path;
            }

            return args;
        }

        private static string ColumnXml(string title)
        {
            return TypedXml("column_clustered", title);
        }

        private static string TypedXml(string type, string title)
        {
            return
                "<ChartConfig><Chart><Type>" + type + "</Type><Title>" + title + "</Title>"
                + "<Width>400</Width><Height>240</Height><DataTable><Header>"
                + "<Column type=\"category\">年</Column>"
                + "<Column type=\"value\">销量</Column>"
                + "</Header><Rows>"
                + "<Row>2023,10</Row><Row>2024,20</Row><Row>2025,30</Row>"
                + "</Rows></DataTable></Chart></ChartConfig>";
        }

        private static string DataXml(string a, string b, string c)
        {
            return
                "<ChartConfig><Chart><Type>column_clustered</Type><Title>" + TitleA + "</Title>"
                + "<DataTable><Header>"
                + "<Column type=\"category\">年</Column>"
                + "<Column type=\"value\">销量</Column>"
                + "</Header><Rows>"
                + "<Row>2023," + a + "</Row><Row>2024," + b + "</Row><Row>2025," + c + "</Row>"
                + "</Rows></DataTable></Chart></ChartConfig>";
        }

        private static string MismatchXml()
        {
            return
                "<ChartConfig><Chart><Type>column_clustered</Type><Title>" + TitleA + "</Title>"
                + "<DataTable><Header>"
                + "<Column type=\"category\">年</Column>"
                + "<Column type=\"value\">销量</Column>"
                + "<Column type=\"value\">利润</Column>"
                + "</Header><Rows>"
                + "<Row>2023,10,1</Row><Row>2024,20,2</Row><Row>2025,30,3</Row>"
                + "</Rows></DataTable></Chart></ChartConfig>";
        }

        private static string FormatXml(string title)
        {
            return "<ChartConfig><Chart><Type>column_clustered</Type><Title>" + title + "</Title></Chart></ChartConfig>";
        }

        private static string WriteXml(string xml)
        {
            string path = Path.Combine(Path.GetTempPath(), "ew-wh-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(path, xml ?? "");
            return path;
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

        private static bool HasChart(WdSession s, string id)
        {
            return Regex.IsMatch(ReadBody(s, id), @"<Chart\s+id=""C_", RegexOptions.IgnoreCase);
        }

        private static bool HasChartId(WdSession s, string id, string chartId)
        {
            if (string.IsNullOrEmpty(chartId))
            {
                return false;
            }

            return ReadBody(s, id).IndexOf("id=\"" + chartId + "\"", StringComparison.Ordinal) >= 0;
        }

        private static int ChartCount(WdSession s, string id)
        {
            return Regex.Matches(ReadBody(s, id), @"<Chart\s+id=""C_", RegexOptions.IgnoreCase).Count;
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
            if (r == null || r.Data == null || string.IsNullOrEmpty(needle))
            {
                return false;
            }

            try
            {
                return JsonConvert.SerializeObject(r.Data).IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
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
