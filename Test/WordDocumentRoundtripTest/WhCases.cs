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
    /// Word 与 WPS 正例都须成功，同一套 ID，不 SKIP。契约负例两侧同一断言。
    /// </summary>
    internal static class WhCases
    {
        public const int Target = 265;

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
            bool wpsWinBefore,
            string part = null)
        {
            bool runFrom136 = string.Equals(part, "from136", StringComparison.OrdinalIgnoreCase);
            bool runBase = !runFrom136 && (string.IsNullOrEmpty(part) || part == "base");
            bool runCon = !runFrom136 && (string.IsNullOrEmpty(part) || part == "con");
            bool runC2 = !runFrom136 && (string.IsNullOrEmpty(part) || part == "c2");
            bool runAttr = !runFrom136 && (string.IsNullOrEmpty(part) || part == "attr");
            var s = new WdSession(host);
            s.DismissOrphanExcelOnClose = !s.IsWps;
            if (runFrom136)
            {
                run.NotePassed(135);
            }
            if (runBase)
            {
                RunNegEarly(run, s);
            }
            string warmPath = WarmWpsChannel(s);
            try
            {
                if (runBase)
                {
                    RunCreateAndMutate(run, s);
                }
                if (runCon)
                {
                    RunConsolidate(run, s);
                }
                if (runC2 || runFrom136)
                {
                    if (runFrom136)
                    {
                        RunConFlow(run, s);
                    }

                    RunConsolidate2(run, s);
                }
                if (runAttr)
                {
                    RunAttrs(run, s);
                }
                if (runBase || runFrom136)
                {
                    RunNegLate(run, s);
                    RunNoQuit(run, wordBefore, wpsBefore, wordWinBefore, wpsWinBefore);
                }
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
                s.CreateChartFromXml(CreateArgs(null, WriteXml("<NotChart><Chart><Type>column</Type></Chart></NotChart>"))),
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
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wh-ctx-" + Guid.NewGuid().ToString("N") + ".xml"),
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

        // ============ 巩固 100（wh-con-*）============

        private const string TitleC = "易写图再改";
        private const string ConSent = "易写图巩固句。";
        private const string PairSent = "易写图巩固乙句。";
        private const string TableCellText = "易写表巩固甲";

        private static readonly string[] ConNegIds =
        {
            "wh-con-n-root-case", "wh-con-n-no-chart", "wh-con-n-type-case",
            "wh-con-n-area", "wh-con-n-primarytitle",
            "wh-con-n-no-header", "wh-con-n-empty-header",
            "wh-con-n-no-value", "wh-con-n-no-rows", "wh-con-n-empty-rows", "wh-con-n-df-noattr",
            "wh-con-n-csv-missing", "wh-con-n-csv-empty", "wh-con-n-axisy-bad", "wh-con-n-broken-xml",
            "wh-con-n-format-only", "wh-con-n-row-short"
        };

        private static readonly string[] ConCreateIds =
        {
            "wh-con-x-no-title", "wh-con-x-width", "wh-con-x-height", "wh-con-x-two-series",
            "wh-con-x-three-series", "wh-con-x-ten-rows", "wh-con-x-one-row", "wh-con-x-decimal",
            "wh-con-x-negative", "wh-con-x-big", "wh-con-x-empty-y", "wh-con-x-nan-y",
            "wh-con-x-fill-off", "wh-con-x-axisx", "wh-con-x-axisy", "wh-con-x-legend",
            "wh-con-x-datalabels", "wh-con-x-color", "wh-con-x-csv", "wh-con-x-three-charts",
            "wh-con-x-bar", "wh-con-x-pie2d",
            "wh-con-p-zero", "wh-con-p-big", "wh-con-p-nonint", "wh-con-p-keep-sent",
            "wh-con-p-xml-decl", "wh-con-p-bom", "wh-con-p-title-escape", "wh-con-p-title-cn",
            "wh-con-p-theme"
        };

        private static readonly string[] ConExtractIds =
        {
            "wh-con-e-type", "wh-con-e-title", "wh-con-e-row-val", "wh-con-e-col-val",
            "wh-con-e-file-content", "wh-con-e-file-nonempty", "wh-con-e-nofile", "wh-con-e-keep-body",
            "wh-con-e-keep-chart", "wh-con-e-second", "wh-con-e-after-fmt", "wh-con-e-after-data",
            "wh-con-e-scope-case", "wh-con-e-unk-channel", "wh-con-e-sel-string"
        };

        private static readonly string[] ConApplyIds =
        {
            "wh-con-a-path-fmt", "wh-con-a-path-data", "wh-con-a-scope-all", "wh-con-a-scope-case",
            "wh-con-a-mode-bad", "wh-con-a-title2", "wh-con-a-data-vals", "wh-con-a-data-keep-body",
            "wh-con-a-data-keep-chart", "wh-con-a-other-chart", "wh-con-a-mis-row-less", "wh-con-a-mis-row-more",
            "wh-con-a-mis-col-less", "wh-con-a-mis-no-write", "wh-con-a-bad-xml", "wh-con-a-root",
            "wh-con-a-unk-channel", "wh-con-a-empty-id", "wh-con-a-data-nodt", "wh-con-a-echo-fields"
        };

        private static readonly string[] ConDeleteIds =
        {
            "wh-con-d-mid", "wh-con-d-all", "wh-con-d-rebuild3", "wh-con-d-unk-channel",
            "wh-con-d-empty-id", "wh-con-d-extract-gone", "wh-con-d-apply-gone", "wh-con-d-rebuild-extract",
            "wh-con-d-rebuild-apply", "wh-con-d-mix-table", "wh-con-d-order", "wh-con-d-sent-keep",
            "wh-con-d-rebuild-type"
        };

        private static readonly string[] ConFlowIds =
        {
            "wh-con-f-cycle", "wh-con-f-cycle-body", "wh-con-f-cycle-clean", "wh-con-f-two-independent",
            "wh-con-f-count"
        };

        private static void RunConsolidate(TestRun run, WdSession s)
        {
            RunConNeg(run, s);
            RunConCreate(run, s);
            RunConExtract(run, s);
            RunConApply(run, s);
            RunConDelete(run, s);
            RunConFlow(run, s);
        }

        /// <summary>建图契约负例 18。契约校验在渠道解析之前，Word / WPS 同一断言。</summary>
        private static void RunConNeg(TestRun run, WdSession s)
        {
            ExpectFail(run, "wh-con-n-root-case",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<chartconfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></chartconfig>"))),
                "ChartConfig");
            ExpectFail(run, "wh-con-n-no-chart",
                s.CreateChartFromXml(CreateArgs(null, WriteXml("<ChartConfig></ChartConfig>"))),
                "Chart");
            ExpectFail(run, "wh-con-n-type-case",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("LINE", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-area",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("area", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-primarytitle",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    GenXml("column", TitleA, V1, R3, "<AxisY><PrimaryTitle>万台</PrimaryTitle></AxisY>")))),
                "须改成");
            ExpectFail(run, "wh-con-n-no-header",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Header");
            ExpectFail(run, "wh-con-n-empty-header",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header></Header>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Column");
            ExpectFail(run, "wh-con-n-no-value",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column></Header>"
                    + "<Rows><Row>2023</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "value");
            ExpectFail(run, "wh-con-n-no-rows",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "</DataTable></Chart></ChartConfig>"))),
                "Rows");
            ExpectFail(run, "wh-con-n-empty-rows",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows></Rows></DataTable></Chart></ChartConfig>"))),
                "Row");
            ExpectFail(run, "wh-con-n-df-noattr",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<DataFile/></DataTable></Chart></ChartConfig>"))),
                "fileName");

            string csvMissing = Path.Combine(Path.GetTempPath(), "ew-wh-csv-no-" + Guid.NewGuid().ToString("N") + ".csv");
            ToolResult rCsvMiss = s.CreateChartFromXml(CreateArgs(null, WriteXml(CsvXml(csvMissing))));
            Ok(run, "wh-con-n-csv-missing",
                rCsvMiss != null && !rCsvMiss.Success && !string.IsNullOrEmpty(rCsvMiss.Error),
                "CSV 不存在应失败且有文案");

            string csvEmpty = Path.Combine(Path.GetTempPath(), "ew-wh-csv-e-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(csvEmpty, "年,销量\r\n");
            ExpectFail(run, "wh-con-n-csv-empty",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(CsvXml(csvEmpty)))),
                "数据行");

            ExpectFail(run, "wh-con-n-axisy-bad",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    GenXml("column", TitleA, V1, R3, null, " axisY=\"middle\"", null)))),
                "AxisY");
            ExpectFail(run, "wh-con-n-broken-xml",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type>"))),
                "XML");
            ExpectFail(run, "wh-con-n-format-only",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(FormatXml(TitleA)))),
                "SeriesCollection");
            ExpectFail(run, "wh-con-n-row-short",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    GenXml("column", TitleA, V1, new[] { "2023" })))),
                "Point");
        }

        /// <summary>建图特性 20 + position/杂项 9。一特性一稿，DataTable 行值互不相同。</summary>
        private static void RunConCreate(TestRun run, WdSession s)
        {
            AssertCreateSolo(run, s, "wh-con-x-no-title",
                GenXml("column", null, V1, Rows("101")));
            AssertCreateSolo(run, s, "wh-con-x-width",
                GenXml("column", TitleA, V1, Rows("102"), "<Width>500</Width>"));
            AssertCreateSolo(run, s, "wh-con-x-height",
                GenXml("column", TitleA, V1, Rows("103"), "<Height>300</Height>"));
            AssertCreateSolo(run, s, "wh-con-x-two-series",
                GenXml("column", TitleA, V2, new[] { "2023,10,1", "2024,20,2", "2025,30,3" }));
            AssertCreateSolo(run, s, "wh-con-x-three-series",
                GenXml("column", TitleA, V3, new[] { "2023,10,1,5", "2024,20,2,6", "2025,30,3,7" }));
            AssertCreateSolo(run, s, "wh-con-x-ten-rows",
                GenXml("column", TitleA, V1, R10));
            AssertCreateSolo(run, s, "wh-con-x-one-row",
                GenXml("column", TitleA, V1, new[] { "2023,104" }));
            AssertCreateSolo(run, s, "wh-con-x-decimal",
                GenXml("column", TitleA, V1, new[] { "2023,3.14", "2024,2.71", "2025,1.41" }));
            AssertCreateSolo(run, s, "wh-con-x-negative",
                GenXml("column", TitleA, V1, new[] { "2023,-5", "2024,20", "2025,-30" }));
            AssertCreateSolo(run, s, "wh-con-x-big",
                GenXml("column", TitleA, V1, new[] { "2023,1234567", "2024,20", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-empty-y",
                GenXml("column", TitleA, V1, new[] { "2023,105", "2024,", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-nan-y",
                GenXml("column", TitleA, V1, new[] { "2023,106", "2024,abc", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-fill-off",
                GenXml("column", TitleA, V1, new[] { "2023,107", "2024,", "2025,30" },
                    "<FillMissingValues>false</FillMissingValues>"));
            AssertCreateSolo(run, s, "wh-con-x-axisx",
                GenXml("column", TitleA, V1, Rows("108"), "<AxisX><Title>年份</Title></AxisX>"));
            AssertCreateSolo(run, s, "wh-con-x-axisy",
                GenXml("column", TitleA, V1, Rows("109"), "<AxisY><Title>万台</Title></AxisY>"));
            AssertCreateSolo(run, s, "wh-con-x-legend",
                GenXml("column", TitleA, V1, Rows("110"), "<LegendPosition>bottom</LegendPosition>"));
            AssertCreateSolo(run, s, "wh-con-x-datalabels",
                GenXml("column", TitleA, V1, Rows("111"), "<ShowDataLabels>true</ShowDataLabels>"));
            AssertCreateSolo(run, s, "wh-con-x-color",
                GenXml("column", TitleA, V1, Rows("112"), null, " color=\"#FF0000\"", null));
            AssertCreateSolo(run, s, "wh-con-x-bar",
                GenXml("bar", TitleA, V1, Rows("114")));
            AssertCreateSolo(run, s, "wh-con-x-pie2d",
                GenXml("pie2d", TitleA, V1, Rows("115")));

            string csvOk = Path.Combine(Path.GetTempPath(), "ew-wh-csv-" + Guid.NewGuid().ToString("N") + ".csv");
            File.WriteAllText(csvOk, "年,销量\r\n2023,113\r\n2024,20\r\n2025,30\r\n");
            AssertCreateSolo(run, s, "wh-con-x-csv", CsvXml(csvOk));

            string path3, id3;
            if (TryOpen(run, s, out path3, out id3, "wh-con-x-three-charts"))
            {
                try
                {
                    if (!s.TypePlain(path3, ConSent))
                    {
                        run.CaseFail("wh-con-x-three-charts", "COM 写巩固句失败");
                    }
                    else
                    {
                        s.ActivateDoc(path3);
                        s.GetContent(id3);
                        string err;
                        bool threeOk = CreateOk(s, id3, GenXml("column", TitleA, V1, Rows("201")), null, out err);
                        s.MoveSelectionToEnd(path3);
                        threeOk = threeOk && CreateOk(s, id3, GenXml("line", TitleLine, V1, Rows("202")), null, out err);
                        s.MoveSelectionToEnd(path3);
                        threeOk = threeOk && CreateOk(s, id3, GenXml("pie3d", TitlePie, V1, Rows("203")), null, out err);
                        Ok(run, "wh-con-x-three-charts",
                            threeOk && ChartCount(s, id3) >= 3,
                            threeOk ? "正文不足 3 张 Chart" : err);
                    }
                }
                finally
                {
                    s.CloseDocByPath(path3);
                }
            }

            AssertCreateSolo(run, s, "wh-con-p-zero",
                GenXml("column", TitleA, V1, Rows("301")), 0);
            AssertCreateSolo(run, s, "wh-con-p-big",
                GenXml("column", TitleA, V1, Rows("302")), 99999);
            AssertCreateSoloRaw(run, s, "wh-con-p-nonint",
                GenXml("column", TitleA, V1, Rows("303")), "abc");

            string pathKeep, idKeep;
            if (TryOpen(run, s, out pathKeep, out idKeep, "wh-con-p-keep-sent"))
            {
                try
                {
                    if (!s.TypePlain(pathKeep, ConSent))
                    {
                        run.CaseFail("wh-con-p-keep-sent", "COM 写巩固句失败");
                    }
                    else
                    {
                        s.ActivateDoc(pathKeep);
                        s.GetContent(idKeep);
                        string err;
                        bool keepOk = CreateOk(s, idKeep, GenXml("column", TitleA, V1, Rows("304")), 0, out err);
                        Ok(run, "wh-con-p-keep-sent",
                            keepOk && StillHas(s, idKeep, ConSent),
                            keepOk ? "position=0 建图丢了巩固句" : err);
                    }
                }
                finally
                {
                    s.CloseDocByPath(pathKeep);
                }
            }

            AssertCreateSoloFile(run, s, "wh-con-p-xml-decl",
                WriteXml("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                    + GenXml("column", TitleA, V1, Rows("305"))));

            string bomPath = Path.Combine(Path.GetTempPath(), "ew-wh-bom-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(bomPath, GenXml("column", TitleA, V1, Rows("306")), new System.Text.UTF8Encoding(true));
            AssertCreateSoloFile(run, s, "wh-con-p-bom", bomPath);

            AssertCreateSolo(run, s, "wh-con-p-title-escape",
                GenXml("column", "易写图&amp;销量", V1, Rows("307")));
            AssertCreateSolo(run, s, "wh-con-p-title-cn",
                GenXml("column", "易写图「销量」（2025）", V1, Rows("308")));
            AssertCreateSolo(run, s, "wh-con-p-theme",
                GenXml("column", TitleA, V1, Rows("309"), "<Theme>Office</Theme>"));
        }

        /// <summary>抽图巩固 15。</summary>
        private static void RunConExtract(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ConExtractIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConSent))
            {
                FailMany(run, "COM 写巩固句失败", ConExtractIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("401")))));
            s.MoveSelectionToEnd(path);
            ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("line", TitleLine, V1, Rows("402")))));
            string chart1 = WdSession.Str(c1, "chart_id");
            string chart2 = WdSession.Str(c2, "chart_id");
            if (string.IsNullOrEmpty(chart1) || string.IsNullOrEmpty(chart2))
            {
                FailMany(run, "建夹具图失败: " + (c1.Error ?? c2.Error ?? ""), ConExtractIds);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult xall = s.ExtractChartXml(ExtractArgs(id, chart1, null, "all"));
            string allXml = WdSession.Str(xall, "xml_content");
            Ok(run, "wh-con-e-type", xall.Success && Contains(allXml, "Type"), xall.Error ?? "抽 all 无 Type");
            Ok(run, "wh-con-e-title", xall.Success && Contains(allXml, TitleA), xall.Error ?? "抽 all 无 Title");

            ToolResult xdata = s.ExtractChartXml(ExtractArgs(id, chart1, null, "data"));
            string dataXml = WdSession.Str(xdata, "xml_content");
            Ok(run, "wh-con-e-row-val", xdata.Success && Contains(dataXml, "2023"), xdata.Error ?? "抽 data 无行值");
            Ok(run, "wh-con-e-col-val", xdata.Success && Contains(dataXml, "销量"), xdata.Error ?? "抽 data 无列名");

            string filePath = Path.Combine(Path.GetTempPath(), "ew-wh-xf-" + Guid.NewGuid().ToString("N") + ".xml");
            ToolResult xfile = s.ExtractChartXml(ExtractArgs(id, chart1, filePath, "all"));
            Ok(run, "wh-con-e-file-content",
                xfile.Success && File.Exists(filePath) && Contains(File.ReadAllText(filePath), "ChartConfig"),
                xfile.Error ?? "落盘文件无 ChartConfig");
            Ok(run, "wh-con-e-file-nonempty",
                xfile.Success && File.Exists(filePath) && new FileInfo(filePath).Length > 10,
                xfile.Error ?? "落盘文件为空");

            ToolResult xnofile = s.ExtractChartXml(ExtractArgsNoPath(id, chart1, "format"));
            Ok(run, "wh-con-e-nofile",
                xnofile.Success && !string.IsNullOrEmpty(WdSession.Str(xnofile, "xml_content")),
                xnofile.Error ?? "无 path 抽图失败");

            Ok(run, "wh-con-e-keep-body", StillHas(s, id, ConSent), "抽图丢了巩固句");
            Ok(run, "wh-con-e-keep-chart", HasChartId(s, id, chart1), "抽图后 C_ 不在正文");

            ToolResult xsecond = s.ExtractChartXml(ExtractArgs(id, chart2, null, "all"));
            Ok(run, "wh-con-e-second",
                xsecond.Success && Contains(WdSession.Str(xsecond, "xml_content"), TitleLine),
                xsecond.Error ?? "第二张抽图失败");

            ToolResult fmt = s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", FormatXml(TitleB)));
            ToolResult xaf = s.ExtractChartXml(ExtractArgs(id, chart1, null, "format"));
            Ok(run, "wh-con-e-after-fmt",
                fmt.Success && xaf.Success && Contains(WdSession.Str(xaf, "xml_content"), TitleB),
                (fmt.Error ?? xaf.Error) ?? "套后抽 format 无新 Title");

            ToolResult dat = s.ApplyChartFromXml(ApplyArgs(id, chart1, "data", DataXml("97", "98", "99")));
            ToolResult xad = s.ExtractChartXml(ExtractArgs(id, chart1, null, "data"));
            Ok(run, "wh-con-e-after-data",
                dat.Success && xad.Success && Contains(WdSession.Str(xad, "xml_content"), "97"),
                (dat.Error ?? xad.Error) ?? "套后抽 data 无新值");

            ToolResult xcase = s.ExtractChartXml(ExtractArgs(id, chart1, null, "DATA"));
            Ok(run, "wh-con-e-scope-case", xcase.Success, xcase.Error ?? "scope 大写失败");

            ExpectFail(run, "wh-con-e-unk-channel",
                s.ExtractChartXml(ExtractArgs("w999", chart1, null, "all")),
                "未知");
            ExpectFail(run, "wh-con-e-sel-string",
                s.ExtractChartXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["chart_selection"] = "by_id",
                    ["path"] = Path.Combine(Path.GetTempPath(), "ew-wh-ss-" + Guid.NewGuid().ToString("N") + ".xml")
                }),
                "chart_selection");

            s.CloseDocByPath(path);
        }

        /// <summary>套用巩固 20。</summary>
        private static void RunConApply(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ConApplyIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConSent))
            {
                FailMany(run, "COM 写巩固句失败", ConApplyIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("501")))));
            s.MoveSelectionToEnd(path);
            ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", "易写图巩固乙", V1, Rows("502")))));
            s.MoveSelectionToEnd(path);
            ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V2, new[] { "2023,10,1", "2024,20,2", "2025,30,3" }))));
            string chart1 = WdSession.Str(c1, "chart_id");
            string chart2 = WdSession.Str(c2, "chart_id");
            string chart3 = WdSession.Str(c3, "chart_id");
            if (string.IsNullOrEmpty(chart1) || string.IsNullOrEmpty(chart2) || string.IsNullOrEmpty(chart3))
            {
                FailMany(run, "建夹具图失败", ConApplyIds);
                s.CloseDocByPath(path);
                return;
            }

            string fmtPath = WriteXml(FormatXml(TitleB));
            ToolResult apath = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = chart1,
                ["scope"] = "format",
                ["path"] = fmtPath
            });
            Ok(run, "wh-con-a-path-fmt", apath.Success, apath.Error ?? "path 套格式失败");

            string dataPath = WriteXml(DataXml("11", "22", "33"));
            ToolResult apathd = s.ApplyChartFromXml(new Dictionary<string, object>
            {
                ["channel_id"] = id,
                ["chart_id"] = chart1,
                ["scope"] = "data",
                ["path"] = dataPath
            });
            Ok(run, "wh-con-a-path-data", apathd.Success, apathd.Error ?? "path 套数据失败");

            ToolResult aall = s.ApplyChartFromXml(ApplyArgs(id, chart1, "all", DataXml("11", "22", "33")));
            Ok(run, "wh-con-a-scope-all", aall.Success, aall.Error ?? "scope=all 失败");

            ToolResult acase = s.ApplyChartFromXml(ApplyArgs(id, chart1, "FORMAT", FormatXml(TitleB)));
            Ok(run, "wh-con-a-scope-case", acase.Success, acase.Error ?? "scope 大写失败");

            ExpectFail(run, "wh-con-a-mode-bad",
                s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", FormatXml(TitleB), "partial")),
                "mode");

            ToolResult aother = s.ApplyChartFromXml(ApplyArgs(id, chart2, "format", FormatXml(TitleC)));
            ToolResult xc1 = s.ExtractChartXml(ExtractArgs(id, chart1, null, "format"));
            Ok(run, "wh-con-a-other-chart",
                aother.Success && xc1.Success && Contains(WdSession.Str(xc1, "xml_content"), TitleB),
                (aother.Error ?? xc1.Error) ?? "套第二张动了第一张");

            ToolResult at1 = s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", FormatXml(TitleB)));
            ToolResult at2 = s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", FormatXml(TitleC)));
            ToolResult xt2 = s.ExtractChartXml(ExtractArgs(id, chart1, null, "format"));
            Ok(run, "wh-con-a-title2",
                at1.Success && at2.Success && xt2.Success && Contains(WdSession.Str(xt2, "xml_content"), TitleC),
                (at1.Error ?? at2.Error ?? xt2.Error) ?? "连套两次 Title 不对");

            ToolResult adv = s.ApplyChartFromXml(ApplyArgs(id, chart1, "data", DataXml("97", "98", "99")));
            ToolResult xdv = s.ExtractChartXml(ExtractArgs(id, chart1, null, "data"));
            string dvXml = WdSession.Str(xdv, "xml_content");
            Ok(run, "wh-con-a-data-vals",
                adv.Success && xdv.Success && Contains(dvXml, "97") && Contains(dvXml, "98") && Contains(dvXml, "99"),
                (adv.Error ?? xdv.Error) ?? "套 data 值不对");
            Ok(run, "wh-con-a-data-keep-body", adv.Success && StillHas(s, id, ConSent), "套 data 丢了巩固句");
            Ok(run, "wh-con-a-data-keep-chart", adv.Success && HasChartId(s, id, chart1), "套 data 后 C_ 不在");

            ExpectFail(run, "wh-con-a-mis-row-less",
                s.ApplyChartFromXml(ApplyArgs(id, chart2, "data", DataXml2("1", "2"))),
                "chart_data_structure_mismatch");
            ExpectFail(run, "wh-con-a-mis-row-more",
                s.ApplyChartFromXml(ApplyArgs(id, chart2, "data", DataXml4("1", "2", "3", "4"))),
                "chart_data_structure_mismatch");
            ExpectFail(run, "wh-con-a-mis-col-less",
                s.ApplyChartFromXml(ApplyArgs(id, chart3, "data", DataXml("1", "2", "3"))),
                "chart_data_structure_mismatch");

            ToolResult xnw = s.ExtractChartXml(ExtractArgs(id, chart2, null, "data"));
            Ok(run, "wh-con-a-mis-no-write",
                xnw.Success && Contains(WdSession.Str(xnw, "xml_content"), "502"),
                xnw.Error ?? "mismatch 后旧值被改");

            ExpectFail(run, "wh-con-a-bad-xml",
                s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", "<ChartConfig><Chart>")),
                "XML");
            ExpectFail(run, "wh-con-a-root",
                s.ApplyChartFromXml(ApplyArgs(id, chart1, "format", "<NotChart></NotChart>")),
                "ChartConfig");
            ExpectFail(run, "wh-con-a-unk-channel",
                s.ApplyChartFromXml(ApplyArgs("w999", chart1, "format", FormatXml(TitleB))),
                "未知");
            ExpectFail(run, "wh-con-a-empty-id",
                s.ApplyChartFromXml(ApplyArgs(id, "", "format", FormatXml(TitleB))),
                "编号");
            ExpectFail(run, "wh-con-a-data-nodt",
                s.ApplyChartFromXml(ApplyArgs(id, chart1, "data", FormatXml(TitleB))),
                "DataTable");

            ToolResult aecho = s.ApplyChartFromXml(ApplyArgs(id, chart1, "data", DataXml("5", "6", "7")));
            Ok(run, "wh-con-a-echo-fields",
                aecho.Success && WdSession.Bool(aecho, "chart_applied") && WdSession.HasKey(aecho, "series_updated"),
                aecho.Error ?? "回包缺 chart_applied / series_updated");

            s.CloseDocByPath(path);
        }

        /// <summary>删 / 重建巩固 13。</summary>
        private static void RunConDelete(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ConDeleteIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConSent + "\r" + PairSent)
                || !s.InsertTable(path, 2, 2, new[] { TableCellText, "乙", "丙", "丁" }))
            {
                FailMany(run, "COM 铺稿失败", ConDeleteIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("601")))));
            s.MoveSelectionToEnd(path);
            ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("602")))));
            s.MoveSelectionToEnd(path);
            ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("603")))));
            string chart1 = WdSession.Str(c1, "chart_id");
            string chart2 = WdSession.Str(c2, "chart_id");
            string chart3 = WdSession.Str(c3, "chart_id");
            if (string.IsNullOrEmpty(chart1) || string.IsNullOrEmpty(chart2) || string.IsNullOrEmpty(chart3))
            {
                FailMany(run, "建夹具图失败", ConDeleteIds);
                s.CloseDocByPath(path);
                return;
            }

            ToolResult dmid = s.DeleteChart(DelArgs(id, chart2));
            Ok(run, "wh-con-d-mid", dmid.Success && ChartCount(s, id) == 2, dmid.Error ?? "删中间图失败");
            Ok(run, "wh-con-d-mix-table", StillHas(s, id, TableCellText), "删图带走了表格");
            Ok(run, "wh-con-d-order", HasChartId(s, id, chart3), "删中间图后第三张 C_ 变了");
            ExpectFail(run, "wh-con-d-extract-gone",
                s.ExtractChartXml(ExtractArgs(id, chart2, null, "all")),
                "找不到");
            ExpectFail(run, "wh-con-d-apply-gone",
                s.ApplyChartFromXml(ApplyArgs(id, chart2, "format", FormatXml(TitleB))),
                "找不到");
            ExpectFail(run, "wh-con-d-unk-channel",
                s.DeleteChart(DelArgs("w999", chart1)),
                "未知");
            ExpectFail(run, "wh-con-d-empty-id",
                s.DeleteChart(DelArgs(id, "")),
                "编号");
            Ok(run, "wh-con-d-sent-keep",
                StillHas(s, id, ConSent) && StillHas(s, id, PairSent), "删图丢了巩固句");

            ToolResult d1 = s.DeleteChart(DelArgs(id, chart1));
            ToolResult d3 = s.DeleteChart(DelArgs(id, chart3));
            Ok(run, "wh-con-d-all", d1.Success && d3.Success && ChartCount(s, id) == 0, "全删后正文还有 Chart");

            s.MoveSelectionToEnd(path);
            ToolResult r1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("611")))));
            s.MoveSelectionToEnd(path);
            ToolResult r2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("612")))));
            s.MoveSelectionToEnd(path);
            ToolResult r3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("613")))));
            string n1 = WdSession.Str(r1, "chart_id");
            string n2 = WdSession.Str(r2, "chart_id");
            string n3 = WdSession.Str(r3, "chart_id");
            Ok(run, "wh-con-d-rebuild3",
                r1.Success && r2.Success && r3.Success
                    && n1.StartsWith("C_", StringComparison.Ordinal)
                    && n2.StartsWith("C_", StringComparison.Ordinal)
                    && n3.StartsWith("C_", StringComparison.Ordinal)
                    && n1 != n2 && n2 != n3 && n1 != n3,
                "重建三张 C_ 不唯一");

            ToolResult xr = s.ExtractChartXml(ExtractArgs(id, n1, null, "all"));
            Ok(run, "wh-con-d-rebuild-extract", xr.Success, xr.Error ?? "重建后抽图失败");
            ToolResult ar = s.ApplyChartFromXml(ApplyArgs(id, n2, "format", FormatXml(TitleB)));
            Ok(run, "wh-con-d-rebuild-apply", ar.Success, ar.Error ?? "重建后套图失败");

            ToolResult rl = s.CreateChartFromXml(CreateArgs(id, WriteXml(TypedXml("line", TitleLine))));
            Ok(run, "wh-con-d-rebuild-type",
                rl.Success && Contains(WdSession.Str(rl, "chart_type"), "line"),
                rl.Error ?? "重建折线 type 不对");

            s.CloseDocByPath(path);
        }

        /// <summary>全链路巩固 5。</summary>
        private static void RunConFlow(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, ConFlowIds))
            {
                return;
            }

            if (!s.TypePlain(path, ConSent))
            {
                FailMany(run, "COM 写巩固句失败", ConFlowIds);
                s.CloseDocByPath(path);
                return;
            }

            s.ActivateDoc(path);
            ToolResult fc = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("701")))));
            string fchart = WdSession.Str(fc, "chart_id");
            bool chain = fc.Success && !string.IsNullOrEmpty(fchart);
            chain = chain && s.ExtractChartXml(ExtractArgs(id, fchart, null, "all")).Success;
            chain = chain && s.ApplyChartFromXml(ApplyArgs(id, fchart, "format", FormatXml(TitleC))).Success;
            ToolResult xf = s.ExtractChartXml(ExtractArgs(id, fchart, null, "format"));
            chain = chain && xf.Success && Contains(WdSession.Str(xf, "xml_content"), TitleC);
            chain = chain && s.ApplyChartFromXml(ApplyArgs(id, fchart, "data", DataXml("97", "98", "99"))).Success;
            ToolResult xd = s.ExtractChartXml(ExtractArgs(id, fchart, null, "data"));
            chain = chain && xd.Success && Contains(WdSession.Str(xd, "xml_content"), "97");
            ToolResult fdel = s.DeleteChart(DelArgs(id, fchart));
            chain = chain && fdel.Success;
            Ok(run, "wh-con-f-cycle", chain, "建抽套删全链断链");
            Ok(run, "wh-con-f-cycle-body", StillHas(s, id, ConSent), "全链丢了巩固句");
            Ok(run, "wh-con-f-cycle-clean", ChartCount(s, id) == 0, "全链删后正文还有 Chart");

            s.MoveSelectionToEnd(path);
            ToolResult ca = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("711")))));
            s.MoveSelectionToEnd(path);
            ToolResult cb = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("712")))));
            string caId = WdSession.Str(ca, "chart_id");
            string cbId = WdSession.Str(cb, "chart_id");
            bool twoOk = ca.Success && cb.Success
                && s.ApplyChartFromXml(ApplyArgs(id, caId, "format", FormatXml(TitleB))).Success
                && s.ApplyChartFromXml(ApplyArgs(id, cbId, "format", FormatXml(TitleC))).Success;
            ToolResult xa = s.ExtractChartXml(ExtractArgs(id, caId, null, "format"));
            ToolResult xb = s.ExtractChartXml(ExtractArgs(id, cbId, null, "format"));
            Ok(run, "wh-con-f-two-independent",
                twoOk
                    && Contains(WdSession.Str(xa, "xml_content"), TitleB)
                    && Contains(WdSession.Str(xb, "xml_content"), TitleC),
                "两图独立套用互相污染");
            Ok(run, "wh-con-f-count", ca.Success && cb.Success && ChartCount(s, id) >= 2, "链路 Chart 计数不对");

            s.CloseDocByPath(path);
        }

        // ============ 巩固再 100（wh-c2-*）============

        private static readonly string[] Con2NegIds =
        {
            "wh-c2-n-column", "wh-c2-n-scatter", "wh-c2-n-radar", "wh-c2-n-doughnut", "wh-c2-n-type-ws",
            "wh-c2-n-empty-file", "wh-c2-n-not-xml", "wh-c2-n-dir", "wh-c2-n-type-pad", "wh-c2-n-row-empty",
            "wh-c2-n-apply-ws-xml", "wh-c2-n-apply-not-xml", "wh-c2-n-apply-html", "wh-c2-n-del-ws-id",
            "wh-c2-n-filename-missing", "wh-c2-n-apply-all-bad", "wh-c2-n-type-typo", "wh-c2-n-type-3dpie",
            "wh-c2-n-create-json", "wh-c2-n-type-combo"
        };

        private static readonly string[] Con2TwinIds =
        {
            "wh-c2-twin-ok", "wh-c2-twin-diff", "wh-c2-twin-count", "wh-c2-twin-x1", "wh-c2-twin-x2",
            "wh-c2-twin-a2", "wh-c2-twin-a1-keep", "wh-c2-twin-d1", "wh-c2-twin-x2-after", "wh-c2-twin-count1",
            "wh-c2-twin-body", "wh-c2-twin-token"
        };

        private static readonly string[] Con2IdIds =
        {
            "wh-c2-id-prefix", "wh-c2-id-alnum", "wh-c2-id-after-data", "wh-c2-id-after-data2",
            "wh-c2-id-after-fmt", "wh-c2-id-in-body", "wh-c2-id-x-after-2data", "wh-c2-id-keep-sent",
            "wh-c2-id-fmt-then-data", "wh-c2-id-all-scope", "wh-c2-id-echo-apply", "wh-c2-id-len"
        };

        private static readonly string[] Con2MixIds =
        {
            "wh-c2-mix-after-table", "wh-c2-mix-keep-table", "wh-c2-mix-two-sent", "wh-c2-mix-cn-cat",
            "wh-c2-mix-long-title", "wh-c2-mix-dup-col", "wh-c2-mix-path-file", "wh-c2-mix-line-after",
            "wh-c2-mix-pie-after", "wh-c2-mix-chart-tbl-chart", "wh-c2-mix-del-keep-tbl", "wh-c2-mix-zero-y",
            "wh-c2-mix-five-ser", "wh-c2-mix-echo-title", "wh-c2-mix-xml-file", "wh-c2-mix-keep-pair"
        };

        private static readonly string[] Con2ExtIds =
        {
            "wh-c2-e-scope-omit", "wh-c2-e-flag", "wh-c2-e-info-id", "wh-c2-e-line-type", "wh-c2-e-pie-type",
            "wh-c2-e-sibling", "wh-c2-e-byid-need", "wh-c2-e-byid-empty", "wh-c2-e-by-index", "wh-c2-e-sel-empty",
            "wh-c2-e-save-false", "wh-c2-e-overwrite", "wh-c2-e-all-rows", "wh-c2-e-format-cfg", "wh-c2-e-keep-body"
        };

        private static readonly string[] Con2AppIds =
        {
            "wh-c2-a-filename", "wh-c2-a-fmt-keep-data", "wh-c2-a-data-keep-title", "wh-c2-a-all-both",
            "wh-c2-a-mode-omit", "wh-c2-a-mode-full", "wh-c2-a-zero", "wh-c2-a-neg", "wh-c2-a-dec",
            "wh-c2-a-twice-title", "wh-c2-a-then-del", "wh-c2-a-echo-series", "wh-c2-a-empty-xml",
            "wh-c2-a-ws-id", "wh-c2-a-self"
        };

        private static readonly string[] Con2DelIds =
        {
            "wh-c2-d-first", "wh-c2-d-last", "wh-c2-d-echo", "wh-c2-d-msg", "wh-c2-d-x-remain",
            "wh-c2-d-count", "wh-c2-d-rebuild", "wh-c2-d-nopfx", "wh-c2-d-conly", "wh-c2-d-keep-line"
        };

        private static void RunConsolidate2(TestRun run, WdSession s)
        {
            RunCon2Neg(run, s);
            RunCon2Twin(run, s);
            RunCon2Id(run, s);
            RunCon2Mix(run, s);
            RunCon2Extract(run, s);
            RunCon2Apply(run, s);
            RunCon2Delete(run, s);
        }

        private static void RunCon2Neg(TestRun run, WdSession s)
        {
            ExpectFail(run, "wh-c2-n-column",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("column_clustered", TitleA, V1, R3)))), "须改成");
            ExpectFail(run, "wh-c2-n-scatter",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("scatter", TitleA, V1, R3)))), "无效");
            ExpectFail(run, "wh-c2-n-radar",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("radar", TitleA, V1, R3)))), "无效");
            ExpectFail(run, "wh-c2-n-doughnut",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("doughnut", TitleA, V1, R3)))), "无效");
            ExpectFail(run, "wh-c2-n-type-ws",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("   ", TitleA, V1, R3)))), "Type");
            ExpectFail(run, "wh-c2-n-empty-file",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(""))), "XML");
            ExpectFail(run, "wh-c2-n-not-xml",
                s.CreateChartFromXml(CreateArgs(null, WriteXml("not-xml"))), "XML");
            ExpectFail(run, "wh-c2-n-dir",
                s.CreateChartFromXml(CreateArgs(null, Path.GetTempPath().TrimEnd('\\'))), "不存在");
            ExpectFail(run, "wh-c2-n-type-pad",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml(" column_clustered ", TitleA, V1, R3)))), "须改成");
            ExpectFail(run, "wh-c2-n-row-empty",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows><Row></Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Point");
            ExpectFail(run, "wh-c2-n-apply-ws-xml",
                s.ApplyChartFromXml(ApplyArgs(null, "C_00001", "format", "   ")), "XML");
            ExpectFail(run, "wh-c2-n-apply-not-xml",
                s.ApplyChartFromXml(ApplyArgs(null, "C_00001", "format", "hello")), "XML");
            ExpectFail(run, "wh-c2-n-apply-html",
                s.ApplyChartFromXml(ApplyArgs(null, "C_00001", "format", "<html></html>")), "ChartConfig");
            ExpectFail(run, "wh-c2-n-del-ws-id",
                s.DeleteChart(new Dictionary<string, object> { ["chart_id"] = "   " }), "编号");
            ExpectFail(run, "wh-c2-n-filename-missing",
                s.CreateChartFromXml(CreateArgsFilename(null, Path.Combine(Path.GetTempPath(), "ew-c2-no-" + Guid.NewGuid().ToString("N") + ".xml"))),
                "不存在");
            ExpectFail(run, "wh-c2-n-apply-all-bad",
                s.ApplyChartFromXml(ApplyArgs(null, "C_00001", "all", "<x/>")), "XML");
            ExpectFail(run, "wh-c2-n-type-typo",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("column_clustred", TitleA, V1, R3)))), "无效");
            ExpectFail(run, "wh-c2-n-type-3dpie",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("3dpie", TitleA, V1, R3)))), "无效");
            ExpectFail(run, "wh-c2-n-create-json",
                s.CreateChartFromXml(CreateArgs(null, WriteXml("{\"Chart\":{}}"))), "XML");
            ExpectFail(run, "wh-c2-n-type-combo",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("combo", TitleA, V1, R3)))), "无效");
        }

        /// <summary>同 DataTable 两张图必须两个 C_，删第一张后第二张仍按号能抽。</summary>
        private static void RunCon2Twin(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenLaid(run, s, ConSent, out path, out id, Con2TwinIds))
            {
                return;
            }

            try
            {
                string xml = GenXml("column", TitleA, V1, Rows("801"));
                ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(xml)));
                s.MoveSelectionToEnd(path);
                ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(xml)));
                string a = WdSession.Str(c1, "chart_id");
                string b = WdSession.Str(c2, "chart_id");
                Ok(run, "wh-c2-twin-ok", c1.Success && c2.Success, c2.Error ?? c1.Error ?? "同数据双建失败");
                Ok(run, "wh-c2-twin-diff",
                    c1.Success && c2.Success && a != b && a.StartsWith("C_", StringComparison.Ordinal) && b.StartsWith("C_", StringComparison.Ordinal),
                    "同数据两张应两个 C_");
                Ok(run, "wh-c2-twin-count", ChartCount(s, id) >= 2, "正文不足两张 Chart");
                Ok(run, "wh-c2-twin-token", IdTokenSafe(a) && IdTokenSafe(b), "C_ 含 / 或 +");

                ToolResult x1 = s.ExtractChartXml(ExtractArgs(id, a, null, "format"));
                ToolResult x2 = s.ExtractChartXml(ExtractArgs(id, b, null, "all"));
                Ok(run, "wh-c2-twin-x1", x1.Success && Contains(WdSession.Str(x1, "xml_content"), TitleA), x1.Error ?? "抽第一张失败");
                Ok(run, "wh-c2-twin-x2", x2.Success, x2.Error ?? "抽第二张失败");

                ToolResult a2 = s.ApplyChartFromXml(ApplyArgs(id, b, "format", FormatXml(TitleC)));
                ToolResult x1b = s.ExtractChartXml(ExtractArgs(id, a, null, "format"));
                Ok(run, "wh-c2-twin-a2", a2.Success, a2.Error ?? "套第二张失败");
                Ok(run, "wh-c2-twin-a1-keep",
                    x1b.Success && Contains(WdSession.Str(x1b, "xml_content"), TitleA),
                    x1b.Error ?? "套第二张动了第一张");

                ToolResult d1 = s.DeleteChart(DelArgs(id, a));
                Ok(run, "wh-c2-twin-d1", d1.Success, d1.Error ?? "删第一张失败");
                ToolResult x2b = s.ExtractChartXml(ExtractArgs(id, b, null, "all"));
                Ok(run, "wh-c2-twin-x2-after", x2b.Success, x2b.Error ?? "删第一张后第二张抽不到");
                Ok(run, "wh-c2-twin-count1", ChartCount(s, id) == 1, "删一张后计数不对");
                Ok(run, "wh-c2-twin-body", StillHas(s, id, ConSent), "双图流程丢了巩固句");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        /// <summary>套 data 后 C_ 不换号，token 字母数字。</summary>
        private static void RunCon2Id(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenLaid(run, s, ConSent, out path, out id, Con2IdIds))
            {
                return;
            }

            try
            {
                ToolResult created = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("811")))));
                string chart = WdSession.Str(created, "chart_id");
                if (!created.Success || string.IsNullOrEmpty(chart))
                {
                    FailMany(run, created.Error ?? "建图无 C_", Con2IdIds);
                    return;
                }

                Ok(run, "wh-c2-id-prefix", chart.StartsWith("C_", StringComparison.Ordinal), "号不以 C_ 开头");
                Ok(run, "wh-c2-id-alnum", IdTokenSafe(chart), "C_ token 含 / + =");
                Ok(run, "wh-c2-id-len", chart.Length >= 10, "C_ 过短");

                ToolResult fmt = s.ApplyChartFromXml(ApplyArgs(id, chart, "format", FormatXml(TitleB)));
                Ok(run, "wh-c2-id-after-fmt", fmt.Success && HasChartId(s, id, chart), fmt.Error ?? "套格式换号");

                ToolResult d1 = s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("21", "22", "23")));
                Ok(run, "wh-c2-id-after-data", d1.Success && HasChartId(s, id, chart), d1.Error ?? "套数据换号");
                ToolResult d2 = s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("31", "32", "33")));
                Ok(run, "wh-c2-id-after-data2", d2.Success && HasChartId(s, id, chart), d2.Error ?? "连套数据换号");
                Ok(run, "wh-c2-id-in-body", HasChartId(s, id, chart), "套数据后正文无该 C_");

                ToolResult xd = s.ExtractChartXml(ExtractArgs(id, chart, null, "data"));
                Ok(run, "wh-c2-id-x-after-2data",
                    xd.Success && Contains(WdSession.Str(xd, "xml_content"), "31"),
                    xd.Error ?? "换号后抽不到或值不对");
                Ok(run, "wh-c2-id-keep-sent", StillHas(s, id, ConSent), "编号块丢了巩固句");

                ToolResult fd = s.ApplyChartFromXml(ApplyArgs(id, chart, "format", FormatXml(TitleC)));
                fd = s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("41", "42", "43")));
                Ok(run, "wh-c2-id-fmt-then-data", fd.Success && HasChartId(s, id, chart), fd.Error ?? "先格式后数据换号");

                ToolResult all = s.ApplyChartFromXml(ApplyArgs(id, chart, "all", DataXml("51", "52", "53")));
                Ok(run, "wh-c2-id-all-scope", all.Success && HasChartId(s, id, chart), all.Error ?? "scope=all 换号");
                Ok(run, "wh-c2-id-echo-apply", WdSession.Bool(all, "chart_applied"), "回包无 chart_applied");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCon2Mix(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, Con2MixIds))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, ConSent + "\r" + PairSent)
                    || !s.InsertTable(path, 2, 2, new[] { TableCellText, "乙", "丙", "丁" }))
                {
                    FailMany(run, "COM 铺混合稿失败", Con2MixIds);
                    return;
                }

                s.ActivateDoc(path);
                s.MoveSelectionToEnd(path);
                ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("821")))));
                string chart1 = WdSession.Str(c1, "chart_id");
                Ok(run, "wh-c2-mix-after-table", c1.Success && !string.IsNullOrEmpty(chart1), c1.Error ?? "表后建图失败");
                Ok(run, "wh-c2-mix-keep-table", StillHas(s, id, TableCellText), "建图丢了表");
                Ok(run, "wh-c2-mix-two-sent", StillHas(s, id, ConSent) && StillHas(s, id, PairSent), "建图丢了双句");
                Ok(run, "wh-c2-mix-keep-pair", StillHas(s, id, PairSent), "建图丢了乙句");

                s.MoveSelectionToEnd(path);
                string cnXml =
                    "<ChartConfig><Chart><Type>column</Type><Title>易写图甲乙</Title><DataTable><Header>"
                    + "<Column type=\"category\">类</Column><Column type=\"value\">值</Column></Header><Rows>"
                    + "<Row>甲,11</Row><Row>乙,22</Row><Row>丙,33</Row></Rows></DataTable></Chart></ChartConfig>";
                ToolResult cn = s.CreateChartFromXml(CreateArgs(id, WriteXml(cnXml)));
                Ok(run, "wh-c2-mix-cn-cat", cn.Success, cn.Error ?? "中文类目建图失败");

                s.MoveSelectionToEnd(path);
                string longTitle = "易写图长标题" + new string('甲', 40);
                ToolResult lng = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", longTitle, V1, Rows("822")))));
                Ok(run, "wh-c2-mix-long-title", lng.Success, lng.Error ?? "长标题建图失败");

                s.MoveSelectionToEnd(path);
                ToolResult dup = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, new[] { "销量", "销量" }, new[] { "2023,1,2", "2024,3,4", "2025,5,6" }))));
                Ok(run, "wh-c2-mix-dup-col", dup.Success, dup.Error ?? "同名两列建图失败");

                string both = WriteXml(GenXml("column", TitleA, V1, Rows("823")));
                s.MoveSelectionToEnd(path);
                ToolResult pf = s.CreateChartFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["path"] = both,
                    ["filename"] = both
                });
                Ok(run, "wh-c2-mix-path-file", pf.Success, pf.Error ?? "path+filename 建图失败");
                Ok(run, "wh-c2-mix-echo-title",
                    pf.Success && Contains(WdSession.Str(pf, "title"), TitleA),
                    "回包 title 不对");
                Ok(run, "wh-c2-mix-xml-file",
                    pf.Success && !string.IsNullOrEmpty(WdSession.Str(pf, "xml_file")),
                    "回包无 xml_file");

                s.MoveSelectionToEnd(path);
                ToolResult line = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("line", TitleLine, V1, Rows("824")))));
                Ok(run, "wh-c2-mix-line-after",
                    line.Success && Contains(WdSession.Str(line, "chart_type"), "line"),
                    line.Error ?? "柱后折线失败");
                s.MoveSelectionToEnd(path);
                ToolResult pie = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("pie3d", TitlePie, V1, Rows("825")))));
                Ok(run, "wh-c2-mix-pie-after",
                    pie.Success && Contains(WdSession.Str(pie, "chart_type"), "pie"),
                    pie.Error ?? "折后饼失败");

                s.MoveSelectionToEnd(path);
                if (!s.InsertTable(path, 2, 2, new[] { "戊", "己", "庚", "辛" }))
                {
                    run.CaseFail("wh-c2-mix-chart-tbl-chart", "COM 插第二张表失败");
                }
                else
                {
                    s.MoveSelectionToEnd(path);
                    ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("826")))));
                    Ok(run, "wh-c2-mix-chart-tbl-chart",
                        c3.Success && ChartCount(s, id) >= 2 && StillHas(s, id, TableCellText),
                        c3.Error ?? "图-表-图失败");
                    if (c3.Success)
                    {
                        ToolResult del = s.DeleteChart(DelArgs(id, WdSession.Str(c3, "chart_id")));
                        Ok(run, "wh-c2-mix-del-keep-tbl",
                            del.Success && StillHas(s, id, TableCellText),
                            del.Error ?? "删图带走了表");
                    }
                    else
                    {
                        run.CaseFail("wh-c2-mix-del-keep-tbl", "无第三张图可删");
                    }
                }

                s.MoveSelectionToEnd(path);
                ToolResult zero = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, new[] { "2023,0", "2024,0", "2025,0" }))));
                Ok(run, "wh-c2-mix-zero-y", zero.Success, zero.Error ?? "全 0 建图失败");

                s.MoveSelectionToEnd(path);
                ToolResult five = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml(
                    "column",
                    TitleA,
                    new[] { "a", "b", "c", "d", "e" },
                    new[] { "2023,1,2,3,4,5", "2024,6,7,8,9,10", "2025,11,12,13,14,15" }))));
                Ok(run, "wh-c2-mix-five-ser", five.Success, five.Error ?? "五系列建图失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCon2Extract(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenLaid(run, s, ConSent, out path, out id, Con2ExtIds))
            {
                return;
            }

            try
            {
                ToolResult col = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("831")))));
                s.MoveSelectionToEnd(path);
                ToolResult line = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("line", TitleLine, V1, Rows("832")))));
                s.MoveSelectionToEnd(path);
                ToolResult pie = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("pie3d", TitlePie, V1, Rows("833")))));
                string cCol = WdSession.Str(col, "chart_id");
                string cLine = WdSession.Str(line, "chart_id");
                string cPie = WdSession.Str(pie, "chart_id");
                if (string.IsNullOrEmpty(cCol) || string.IsNullOrEmpty(cLine) || string.IsNullOrEmpty(cPie))
                {
                    FailMany(run, "建抽图夹具失败", Con2ExtIds);
                    return;
                }

                var omit = ExtractArgs(id, cCol, null, "all");
                omit.Remove("scope");
                ToolResult xomit = s.ExtractChartXml(omit);
                Ok(run, "wh-c2-e-scope-omit", xomit.Success, xomit.Error ?? "省略 scope 失败");
                Ok(run, "wh-c2-e-flag", xomit.Success && WdSession.Bool(xomit, "chart_extracted"), "无 chart_extracted");
                Ok(run, "wh-c2-e-info-id", xomit.Success && BlobHas(xomit, cCol), "回包无该 C_");
                Ok(run, "wh-c2-e-all-rows", xomit.Success && Contains(WdSession.Str(xomit, "xml_content"), "Rows"), "all 无 Rows");

                ToolResult xf = s.ExtractChartXml(ExtractArgs(id, cCol, null, "format"));
                Ok(run, "wh-c2-e-format-cfg", xf.Success && Contains(WdSession.Str(xf, "xml_content"), "ChartConfig"), xf.Error ?? "format 无 ChartConfig");

                ToolResult xl = s.ExtractChartXml(ExtractArgs(id, cLine, null, "all"));
                Ok(run, "wh-c2-e-line-type",
                    xl.Success && (Contains(WdSession.Str(xl, "xml_content"), "line") || Contains(WdSession.Str(xl, "chart_type"), "line") || BlobHas(xl, "line")),
                    xl.Error ?? "折线抽图无 line");
                ToolResult xp = s.ExtractChartXml(ExtractArgs(id, cPie, null, "all"));
                Ok(run, "wh-c2-e-pie-type",
                    xp.Success && (Contains(WdSession.Str(xp, "xml_content"), "pie") || BlobHas(xp, "pie")),
                    xp.Error ?? "饼图抽图无 pie");

                ToolResult delLine = s.DeleteChart(DelArgs(id, cLine));
                ToolResult xcol = s.ExtractChartXml(ExtractArgs(id, cCol, null, "all"));
                Ok(run, "wh-c2-e-sibling", delLine.Success && xcol.Success, xcol.Error ?? "删折线后柱形抽不到");

                ExpectFail(run, "wh-c2-e-byid-need",
                    s.ExtractChartXml(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["chart_selection"] = new Dictionary<string, object> { ["type"] = "by_id" },
                        ["path"] = Path.Combine(Path.GetTempPath(), "ew-c2-e1-" + Guid.NewGuid().ToString("N") + ".xml")
                    }),
                    "chart_id");
                ExpectFail(run, "wh-c2-e-byid-empty",
                    s.ExtractChartXml(ExtractArgs(id, "", null, "all")),
                    "chart_id");
                ExpectFail(run, "wh-c2-e-by-index",
                    s.ExtractChartXml(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["chart_selection"] = new Dictionary<string, object> { ["type"] = "by_index", ["index"] = 0 },
                        ["path"] = Path.Combine(Path.GetTempPath(), "ew-c2-e2-" + Guid.NewGuid().ToString("N") + ".xml")
                    }),
                    "by_id");
                ExpectFail(run, "wh-c2-e-sel-empty",
                    s.ExtractChartXml(new Dictionary<string, object>
                    {
                        ["channel_id"] = id,
                        ["chart_selection"] = new Dictionary<string, object>(),
                        ["path"] = Path.Combine(Path.GetTempPath(), "ew-c2-e3-" + Guid.NewGuid().ToString("N") + ".xml")
                    }),
                    "选择");

                ToolResult nosave = s.ExtractChartXml(ExtractArgsNoPath(id, cCol, "format"));
                nosave = s.ExtractChartXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["scope"] = "format",
                    ["save_to_file"] = false,
                    ["chart_selection"] = new Dictionary<string, object> { ["type"] = "by_id", ["chart_id"] = cCol }
                });
                Ok(run, "wh-c2-e-save-false",
                    nosave.Success && !string.IsNullOrEmpty(WdSession.Str(nosave, "xml_content")),
                    nosave.Error ?? "save_to_file=false 失败");

                string fp = Path.Combine(Path.GetTempPath(), "ew-c2-ow-" + Guid.NewGuid().ToString("N") + ".xml");
                ToolResult o1 = s.ExtractChartXml(ExtractArgs(id, cCol, fp, "all"));
                ToolResult o2 = s.ExtractChartXml(ExtractArgs(id, cCol, fp, "data"));
                Ok(run, "wh-c2-e-overwrite",
                    o1.Success && o2.Success && File.Exists(fp),
                    o2.Error ?? "二次落盘失败");
                Ok(run, "wh-c2-e-keep-body", StillHas(s, id, ConSent), "抽图丢了巩固句");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCon2Apply(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenLaid(run, s, ConSent, out path, out id, Con2AppIds))
            {
                return;
            }

            try
            {
                ToolResult created = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("841")))));
                string chart = WdSession.Str(created, "chart_id");
                if (!created.Success || string.IsNullOrEmpty(chart))
                {
                    FailMany(run, created.Error ?? "建套用夹具失败", Con2AppIds);
                    return;
                }

                string fn = WriteXml(FormatXml(TitleB));
                ToolResult afn = s.ApplyChartFromXml(new Dictionary<string, object>
                {
                    ["channel_id"] = id,
                    ["chart_id"] = chart,
                    ["scope"] = "format",
                    ["filename"] = fn
                });
                Ok(run, "wh-c2-a-filename", afn.Success, afn.Error ?? "filename 套格式失败");

                ToolResult adata = s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("61", "62", "63")));
                ToolResult xdata = s.ExtractChartXml(ExtractArgs(id, chart, null, "data"));
                ToolResult xfmt = s.ExtractChartXml(ExtractArgs(id, chart, null, "format"));
                Ok(run, "wh-c2-a-fmt-keep-data",
                    adata.Success && xdata.Success && Contains(WdSession.Str(xdata, "xml_content"), "61"),
                    xdata.Error ?? "套数据后抽不到新值");
                Ok(run, "wh-c2-a-data-keep-title",
                    xfmt.Success && Contains(WdSession.Str(xfmt, "xml_content"), TitleB),
                    xfmt.Error ?? "套数据冲掉了 Title");

                ToolResult aall = s.ApplyChartFromXml(ApplyArgs(id, chart, "all", DataXml("71", "72", "73")));
                ToolResult xall = s.ExtractChartXml(ExtractArgs(id, chart, null, "data"));
                Ok(run, "wh-c2-a-all-both",
                    aall.Success && xall.Success && Contains(WdSession.Str(xall, "xml_content"), "71"),
                    aall.Error ?? "scope=all 失败");

                var omit = ApplyArgs(id, chart, "format", FormatXml(TitleC));
                omit.Remove("mode");
                ToolResult aomit = s.ApplyChartFromXml(omit);
                Ok(run, "wh-c2-a-mode-omit", aomit.Success, aomit.Error ?? "省略 mode 失败");
                ToolResult afull = s.ApplyChartFromXml(ApplyArgs(id, chart, "format", FormatXml(TitleC), "full"));
                Ok(run, "wh-c2-a-mode-full", afull.Success, afull.Error ?? "mode=full 失败");

                Ok(run, "wh-c2-a-zero",
                    s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("0", "0", "0"))).Success,
                    "套 0 失败");
                Ok(run, "wh-c2-a-neg",
                    s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("-1", "-2", "3"))).Success,
                    "套负数失败");
                Ok(run, "wh-c2-a-dec",
                    s.ApplyChartFromXml(ApplyArgs(id, chart, "data", DataXml("1.5", "2.5", "3.5"))).Success,
                    "套小数失败");

                ToolResult t1 = s.ApplyChartFromXml(ApplyArgs(id, chart, "format", FormatXml(TitleB)));
                ToolResult t2 = s.ApplyChartFromXml(ApplyArgs(id, chart, "format", FormatXml(TitleB)));
                Ok(run, "wh-c2-a-twice-title", t1.Success && t2.Success, t2.Error ?? "连套同一 Title 失败");
                Ok(run, "wh-c2-a-echo-series",
                    WdSession.HasKey(adata, "series_updated") || WdSession.HasKey(aall, "series_updated"),
                    "回包无 series_updated");
                Ok(run, "wh-c2-a-self", HasChartId(s, id, chart) && StillHas(s, id, ConSent), "套用后号或句丢了");

                ExpectFail(run, "wh-c2-a-empty-xml",
                    s.ApplyChartFromXml(ApplyArgs(id, chart, "format", "   ")),
                    "XML");
                ExpectFail(run, "wh-c2-a-ws-id",
                    s.ApplyChartFromXml(ApplyArgs(id, "   ", "format", FormatXml(TitleB))),
                    "编号");

                ToolResult del = s.DeleteChart(DelArgs(id, chart));
                Ok(run, "wh-c2-a-then-del", del.Success && !HasChartId(s, id, chart), del.Error ?? "套后删失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void RunCon2Delete(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpenLaid(run, s, ConSent, out path, out id, Con2DelIds))
            {
                return;
            }

            try
            {
                ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("851")))));
                s.MoveSelectionToEnd(path);
                ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("line", TitleLine, V1, Rows("852")))));
                s.MoveSelectionToEnd(path);
                ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("853")))));
                string a = WdSession.Str(c1, "chart_id");
                string b = WdSession.Str(c2, "chart_id");
                string c = WdSession.Str(c3, "chart_id");
                if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b) || string.IsNullOrEmpty(c))
                {
                    FailMany(run, "建删图夹具失败", Con2DelIds);
                    return;
                }

                ToolResult dFirst = s.DeleteChart(DelArgs(id, a));
                Ok(run, "wh-c2-d-first", dFirst.Success && !HasChartId(s, id, a), dFirst.Error ?? "删第一张失败");
                Ok(run, "wh-c2-d-echo", dFirst.Success && WdSession.Bool(dFirst, "chart_deleted"), "回包无 chart_deleted");
                Ok(run, "wh-c2-d-msg", dFirst.Success && Contains(WdSession.Str(dFirst, "message"), "删除"), "回包 message 无删除");
                ToolResult xb = s.ExtractChartXml(ExtractArgs(id, b, null, "all"));
                Ok(run, "wh-c2-d-x-remain", xb.Success, xb.Error ?? "删第一张后中间张抽不到");
                Ok(run, "wh-c2-d-keep-line", HasChartId(s, id, b), "删第一张带走了折线");

                ToolResult dLast = s.DeleteChart(DelArgs(id, c));
                Ok(run, "wh-c2-d-last", dLast.Success && !HasChartId(s, id, c), dLast.Error ?? "删最后一张失败");
                Ok(run, "wh-c2-d-count", ChartCount(s, id) == 1, "删首尾后应剩 1 张");

                ExpectFail(run, "wh-c2-d-nopfx",
                    s.DeleteChart(DelArgs(id, "00001")),
                    "找不到");
                ExpectFail(run, "wh-c2-d-conly",
                    s.DeleteChart(DelArgs(id, "C_")),
                    "找不到");

                s.MoveSelectionToEnd(path);
                ToolResult reb = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column", TitleA, V1, Rows("854")))));
                Ok(run, "wh-c2-d-rebuild",
                    reb.Success && WdSession.Str(reb, "chart_id").StartsWith("C_", StringComparison.Ordinal) && StillHas(s, id, ConSent),
                    reb.Error ?? "删后重建失败");
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static bool TryOpenLaid(
            TestRun run,
            WdSession s,
            string sent,
            out string path,
            out string id,
            params string[] failIds)
        {
            if (!TryOpen(run, s, out path, out id, failIds))
            {
                return false;
            }

            if (!s.TypePlain(path, sent))
            {
                FailMany(run, "COM 写巩固句失败", failIds);
                s.CloseDocByPath(path);
                path = null;
                id = null;
                return false;
            }

            s.ActivateDoc(path);
            s.GetContent(id);
            return true;
        }

        private static bool IdTokenSafe(string chartId)
        {
            if (string.IsNullOrEmpty(chartId) || !chartId.StartsWith("C_", StringComparison.Ordinal))
            {
                return false;
            }

            string token = chartId.Substring(2);
            foreach (char ch in token)
            {
                if (!((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')))
                {
                    return false;
                }
            }

            return token.Length > 0;
        }

        // ============ 属性覆盖（wh-attr-*）：每个支持字段写 → 抽回 → 再套，C_ 不变 ============

        private static readonly string[] AttrIds =
        {
            "wh-attr-n-data",
            "wh-attr-col", "wh-attr-bool",
            "wh-attr-axis-x", "wh-attr-axis-y", "wh-attr-axis-y2", "wh-attr-tick-rest",
            "wh-attr-series", "wh-attr-markers", "wh-attr-labels",
            "wh-attr-type-bar", "wh-attr-type-line", "wh-attr-type-pie2d", "wh-attr-type-pie", "wh-attr-type-pie3d",
            "wh-attr-pie", "wh-attr-pie-pct",
            "wh-attr-legend-top", "wh-attr-legend-left", "wh-attr-legend-right", "wh-attr-legend-none"
        };

        private static void RunAttrs(TestRun run, WdSession s)
        {
            if (AttrIds.Length != 21)
            {
                FailMany(run, "属性用例登记不是 21 条", AttrIds);
                return;
            }

            ExpectFail(run, "wh-attr-n-data",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart data-plot-box=\"0.1,0.2,0.7,0.6\"><Type>column</Type><Title>易写图禁前缀</Title>"
                    + "<DataTable><Header><Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows><Row>2023,1</Row><Row>2024,2</Row><Row>2025,3</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "data-");

            AssertAttrRoundtrip(run, s, "wh-attr-col",
                AttrChart("column", "易写图柱皮",
                    "<Theme>office</Theme><Width>360</Width><Height>240</Height>"
                    + "<ShowDataLabels>false</ShowDataLabels><LegendPosition>bottom</LegendPosition>"
                    + "<PlotColor>#D9E2F3</PlotColor><TitleFontSize>16</TitleFontSize><TitleFontBold>true</TitleFontBold>"
                    + "<TitleFontColor>#1F4E79</TitleFontColor><Gridlines>false</Gridlines><GapWidth>80</GapWidth>"
                    + "<Overlap>20</Overlap><PlotBox>0.12,0.18,0.70,0.62</PlotBox><PlotInside>0.16,0.22,0.60,0.50</PlotInside>"
                    + "<ChartAreaColor>#F2F2F2</ChartAreaColor><ChartStyle>10</ChartStyle><LegendFontColor>#333333</LegendFontColor>"
                    + "<DataMarkers>false</DataMarkers><YAxisMin>10</YAxisMin><YAxisMax>100</YAxisMax><YAxisMajorUnit>10</YAxisMajorUnit>"
                    + "<FillMissingValues>true</FillMissingValues>",
                    Col("", "销量"),
                    Rows3(401)),
                "<Type>column</Type>", "<Title>易写图柱皮</Title>", "<Theme>office</Theme>",
                "<Width>360</Width>", "<Height>240</Height>",
                "<ShowDataLabels>false</ShowDataLabels>", "<LegendPosition>bottom</LegendPosition>",
                "D9E2F3", "<TitleFontSize>16</TitleFontSize>", "<TitleFontBold>true</TitleFontBold>", "1F4E79",
                "<Gridlines>false</Gridlines>", "<GapWidth>80</GapWidth>", "<Overlap>20</Overlap>",
                "0.12,0.18,0.70,0.62", "0.16,0.22,0.60,0.50", "F2F2F2", "<ChartStyle>10</ChartStyle>", "333333",
                "<DataMarkers>false</DataMarkers>", "<YAxisMin>10</YAxisMin>", "<YAxisMax>100</YAxisMax>",
                "<YAxisMajorUnit>10</YAxisMajorUnit>", "<FillMissingValues>true</FillMissingValues>");

            AssertAttrRoundtrip(run, s, "wh-attr-bool",
                AttrChart("column", "易写图布尔",
                    "<ShowDataLabels>true</ShowDataLabels><Gridlines>true</Gridlines>"
                    + "<TitleFontBold>false</TitleFontBold><FillMissingValues>false</FillMissingValues>",
                    Col("", "销量"),
                    Rows3(411)),
                "<ShowDataLabels>true</ShowDataLabels>", "<Gridlines>true</Gridlines>",
                "<TitleFontBold>false</TitleFontBold>", "<FillMissingValues>false</FillMissingValues>");

            AssertAttrRoundtrip(run, s, "wh-attr-axis-x",
                AttrChart("column", "易写图横轴",
                    "<AxisX>" + AxisChrome("年份", "next", "outside", "inside", "true", "true",
                        "<Type>category</Type><TickLabelCount>3</TickLabelCount><TickLabelSpacing>1</TickLabelSpacing><BetweenCategories>true</BetweenCategories>")
                    + "</AxisX>",
                    Col("", "销量"),
                    Rows3(421)),
                AxisNeedles("年份", "next", "outside", "inside", "true", "true",
                    "<Type>category</Type>", "<TickLabelCount>3</TickLabelCount>", "<TickLabelSpacing>1</TickLabelSpacing>", "BetweenCategories"));

            AssertAttrRoundtrip(run, s, "wh-attr-axis-y",
                AttrChart("column", "易写图纵轴",
                    "<AxisY>" + AxisChrome("万台", "low", "inside", "outside", "true", "false") + "</AxisY>",
                    Col("", "销量"),
                    Rows3(431)),
                AxisNeedles("万台", "low", "inside", "outside", "true", "false"));

            AssertAttrRoundtrip(run, s, "wh-attr-axis-y2",
                AttrChart("column", "易写图次轴",
                    "<AxisY2>" + AxisChrome("增速", "high", "cross", "cross", "false", "true") + "</AxisY2>",
                    Col("axisY=\"y\" chartType=\"column\"", "销量") + Col("axisY=\"y2\" chartType=\"line\"", "增速"),
                    "<Row>2023,441,1</Row><Row>2024,442,2</Row><Row>2025,443,3</Row>"),
                AxisNeedles("增速", "high", "cross", "cross", "false", "true"));

            AssertAttrRoundtrip(run, s, "wh-attr-tick-rest",
                AttrChart("column", "易写图刻度",
                    "<AxisX><TickPosition>none</TickPosition><MajorTick>none</MajorTick><MinorTick>none</MinorTick></AxisX>",
                    Col("", "销量"),
                    Rows3(451)),
                "<TickPosition>none</TickPosition>", "<MajorTick>none</MajorTick>", "<MinorTick>none</MinorTick>");

            AssertAttrRoundtrip(run, s, "wh-attr-series",
                AttrChart("line", "易写图系列",
                    null,
                    Col("axisY=\"y2\" color=\"#ED7D31\" chartType=\"line\" showDataLabels=\"true\" showValue=\"true\" showPercentage=\"false\""
                        + " fillGradient=\"0:#FF0000@0;1:#0000FF@0\" fillAngle=\"45\" line=\"#ED7D31\" lineWeight=\"2.5\""
                        + " marker=\"circle\" markerSize=\"8\" markerColor=\"#000000\" markerFill=\"#FFFFFF\""
                        + " labelPosition=\"center\" labelFont=\"微软雅黑\" labelSize=\"12\" labelColor=\"#333333\" labelFormat=\"0.0\"",
                        "增速")
                    + Col("axisY=\"y\" showDataLabels=\"false\" showValue=\"false\" showPercentage=\"true\"", "销量"),
                    "<Row>2023,1,461</Row><Row>2024,2,462</Row><Row>2025,3,463</Row>"),
                "axisY=\"y2\"", "axisY=\"y\"", "ED7D31", "chartType=\"line\"",
                "showDataLabels=\"true\"", "showDataLabels=\"false\"",
                "showValue=\"true\"", "showValue=\"false\"", "showPercentage=\"false\"", "showPercentage=\"true\"",
                "0:#FF0000@0;1:#0000FF@0", "fillAngle=\"45\"", "line=", "lineWeight=\"2.5\"",
                "marker=\"circle\"", "markerSize=\"8\"", "markerColor=\"#000000\"", "markerFill=\"#FFFFFF\"",
                "labelPosition=\"center\"", "labelFont=\"微软雅黑\"", "labelSize=\"12\"", "333333", "labelFormat=\"0.0\"");

            AssertAttrRoundtrip(run, s, "wh-attr-markers",
                AttrChart("line", "易写图标记", null, MarkerColumns(), MarkerRows()),
                "marker=\"none\"", "marker=\"circle\"", "marker=\"diamond\"", "marker=\"square\"",
                "marker=\"triangle\"", "marker=\"star\"", "marker=\"plus\"", "marker=\"x\"", "marker=\"dot\"");

            AssertAttrRoundtrip(run, s, "wh-attr-labels",
                AttrChart("column", "易写图标签位", null, LabelColumns(), LabelRows()),
                "labelPosition=\"center\"", "labelPosition=\"above\"", "labelPosition=\"below\"",
                "labelPosition=\"left\"", "labelPosition=\"right\"", "labelPosition=\"outside\"",
                "labelPosition=\"inside-end\"", "labelPosition=\"inside-base\"");

            AssertAttrRoundtrip(run, s, "wh-attr-type-bar",
                AttrChart("bar", "易写图条形", null, Col("", "销量"), Rows3(471)),
                "<Type>bar</Type>");
            AssertAttrRoundtrip(run, s, "wh-attr-type-line",
                AttrChart("line", "易写图折线皮",
                    "<DataMarkers>true</DataMarkers><MarkerSize>8</MarkerSize><LineWeight>2.25</LineWeight>",
                    Col("", "销量"),
                    Rows3(481)),
                "<Type>line</Type>", "<DataMarkers>true</DataMarkers>", "<MarkerSize>8</MarkerSize>", "<LineWeight>2.25</LineWeight>");
            AssertAttrRoundtrip(run, s, "wh-attr-type-pie2d",
                AttrChart("pie2d", "易写图饼二", null, Col("", "销量"), Rows3(491)),
                "<Type>pie2d</Type>");
            AssertAttrRoundtrip(run, s, "wh-attr-type-pie",
                AttrChart("pie", "易写图饼别名", null, Col("", "销量"), Rows3(501)),
                "<Type>pie2d</Type>");
            AssertAttrRoundtrip(run, s, "wh-attr-type-pie3d",
                AttrChart("pie3d", "易写图饼三", null, Col("", "销量"), Rows3(511)),
                "<Type>pie3d</Type>");

            AssertAttrRoundtrip(run, s, "wh-attr-pie",
                AttrChart("pie3d", "易写图爆炸",
                    "<Explosion>15</Explosion><ShowValue>true</ShowValue><ShowPercentage>false</ShowPercentage>",
                    Col("", "销量"),
                    Rows3(521)),
                "<Type>pie3d</Type>", "<Explosion>15</Explosion>", "<ShowValue>true</ShowValue>", "<ShowPercentage>false</ShowPercentage>");
            AssertAttrRoundtrip(run, s, "wh-attr-pie-pct",
                AttrChart("pie3d", "易写图百分比",
                    "<ShowValue>false</ShowValue><ShowPercentage>true</ShowPercentage>",
                    Col("", "销量"),
                    Rows3(531)),
                "<ShowValue>false</ShowValue>", "<ShowPercentage>true</ShowPercentage>");

            AssertAttrRoundtrip(run, s, "wh-attr-legend-top",
                AttrChart("column", "易写图例上", "<LegendPosition>top</LegendPosition>", Col("", "销量"), Rows3(541)),
                "<LegendPosition>top</LegendPosition>");
            AssertAttrRoundtrip(run, s, "wh-attr-legend-left",
                AttrChart("column", "易写图例左", "<LegendPosition>left</LegendPosition>", Col("", "销量"), Rows3(551)),
                "<LegendPosition>left</LegendPosition>");
            AssertAttrRoundtrip(run, s, "wh-attr-legend-right",
                AttrChart("column", "易写图例右", "<LegendPosition>right</LegendPosition>", Col("", "销量"), Rows3(561)),
                "<LegendPosition>right</LegendPosition>");
            AssertAttrRoundtrip(run, s, "wh-attr-legend-none",
                AttrChart("column", "易写图例无", "<LegendPosition>none</LegendPosition>", Col("", "销量"), Rows3(571)),
                "<LegendPosition>none</LegendPosition>");
        }

        private static string[] AxisNeedles(
            string title,
            string tickPosition,
            string major,
            string minor,
            string visible,
            string grid,
            params string[] extra)
        {
            var list = new List<string>
            {
                "<Title>" + title + "</Title>",
                "<Visible>" + visible + "</Visible>",
                "微软雅黑",
                "C00000",
                "<TickSize>9</TickSize>",
                "<TickPosition>" + tickPosition + "</TickPosition>",
                "<MajorTick>" + major + "</MajorTick>",
                "<MinorTick>" + minor + "</MinorTick>",
                "<Format>0.0</Format>",
                "<Grid>" + grid + "</Grid>",
                "D9D9D9",
                "7F7F7F",
                "<LineWeight>1.5</LineWeight>"
            };
            if (extra != null)
            {
                list.AddRange(extra);
            }

            return list.ToArray();
        }

        private static string AxisChrome(
            string title,
            string tickPosition,
            string major,
            string minor,
            string visible,
            string grid,
            string extra = null)
        {
            return "<Title>" + title + "</Title>"
                + (extra ?? "")
                + "<Visible>" + visible + "</Visible><TickFont>微软雅黑</TickFont><TickColor>#C00000</TickColor><TickSize>9</TickSize>"
                + "<TickPosition>" + tickPosition + "</TickPosition><MajorTick>" + major + "</MajorTick><MinorTick>" + minor + "</MinorTick>"
                + "<Format>0.0</Format><Grid>" + grid + "</Grid><GridColor>#D9D9D9</GridColor><Line>#7F7F7F</Line><LineWeight>1.5</LineWeight>";
        }

        private static string MarkerColumns()
        {
            string[] markers = { "none", "circle", "diamond", "square", "triangle", "star", "plus", "x", "dot" };
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < markers.Length; i++)
            {
                sb.Append(Col("chartType=\"line\" marker=\"" + markers[i] + "\"", "标" + (i + 1)));
            }

            return sb.ToString();
        }

        private static string MarkerRows()
        {
            return "<Row>2023,1,2,3,4,5,6,7,8,9</Row><Row>2024,2,3,4,5,6,7,8,9,10</Row><Row>2025,3,4,5,6,7,8,9,10,11</Row>";
        }

        private static string LabelColumns()
        {
            string[] positions = { "center", "above", "below", "left", "right", "outside", "inside-end", "inside-base" };
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < positions.Length; i++)
            {
                sb.Append(Col("labelPosition=\"" + positions[i] + "\"", "位" + (i + 1)));
            }

            return sb.ToString();
        }

        private static string LabelRows()
        {
            return "<Row>2023,11,12,13,14,15,16,17,18</Row><Row>2024,21,22,23,24,25,26,27,28</Row><Row>2025,31,32,33,34,35,36,37,38</Row>";
        }

        private static string AttrChart(string type, string title, string extras, string columns, string rows)
        {
            return "<ChartConfig><Chart><Type>" + type + "</Type><Title>" + title + "</Title>"
                + (extras ?? "")
                + "<DataTable><Header><Column type=\"category\">年</Column>"
                + columns
                + "</Header><Rows>" + rows + "</Rows></DataTable></Chart></ChartConfig>";
        }

        private static string Col(string attrs, string name)
        {
            return "<Column type=\"value\""
                + (string.IsNullOrEmpty(attrs) ? "" : " " + attrs)
                + ">" + name + "</Column>";
        }

        private static string Rows3(int y0)
        {
            return "<Row>2023," + y0 + "</Row><Row>2024," + (y0 + 1) + "</Row><Row>2025," + (y0 + 2) + "</Row>";
        }

        private static void AssertAttrRoundtrip(TestRun run, WdSession s, string caseId, string xml, params string[] needles)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, ConSent))
                {
                    run.CaseFail(caseId, "COM 写巩固句失败");
                    return;
                }

                s.ActivateDoc(path);
                s.GetContent(id);
                ToolResult created = s.CreateChartFromXml(CreateArgs(id, WriteXml(xml)));
                string chartId = WdSession.Str(created, "chart_id");
                if (created == null || !created.Success || string.IsNullOrEmpty(chartId)
                    || !chartId.StartsWith("C_", StringComparison.Ordinal))
                {
                    run.CaseFail(caseId, created == null ? "空调用结果" : (created.Error ?? "建图失败"));
                    return;
                }

                string missing = FirstMissing(ExtractXml(s, id, chartId), needles);
                if (missing != null)
                {
                    run.CaseFail(caseId, "抽出缺少 " + missing);
                    return;
                }

                ToolResult applied = s.ApplyChartFromXml(ApplyArgs(id, chartId, "all", xml));
                bool idOk = applied != null && applied.Success && HasChartId(s, id, chartId);
                string echoed = applied == null ? "" : WdSession.Str(applied, "chart_id");
                if (idOk && !string.IsNullOrEmpty(echoed) && !string.Equals(echoed, chartId, StringComparison.Ordinal))
                {
                    idOk = false;
                }

                if (!idOk)
                {
                    run.CaseFail(caseId, applied != null && applied.Success ? "再套后 C_ 变了" : (applied == null ? "再套失败" : (applied.Error ?? "再套失败")));
                    return;
                }

                missing = FirstMissing(ExtractXml(s, id, chartId), needles);
                if (missing != null)
                {
                    run.CaseFail(caseId, "再套后抽出缺少 " + missing);
                    return;
                }

                if (!StillHas(s, id, ConSent))
                {
                    run.CaseFail(caseId, "前句丢了");
                    return;
                }

                run.CaseOk(caseId);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static string ExtractXml(WdSession s, string channelId, string chartId)
        {
            ToolResult extracted = s.ExtractChartXml(ExtractArgs(channelId, chartId, null, "all"));
            if (extracted == null || !extracted.Success)
            {
                return "";
            }

            return WdSession.Str(extracted, "xml_content") ?? "";
        }

        private static string FirstMissing(string xml, string[] needles)
        {
            if (string.IsNullOrEmpty(xml))
            {
                return "抽出 XML";
            }

            if (needles == null)
            {
                return null;
            }

            foreach (string needle in needles)
            {
                if (!Contains(xml, needle))
                {
                    return needle;
                }
            }

            return null;
        }

        // ---- 巩固块夹具 ----

        private static readonly string[] V1 = { "销量" };
        private static readonly string[] V2 = { "销量", "利润" };
        private static readonly string[] V3 = { "销量", "利润", "成本" };
        private static readonly string[] R3 = { "2023,10", "2024,20", "2025,30" };
        private static readonly string[] R3S2 = { "2023,10,1", "2024,20,2", "2025,30,3" };
        private static readonly string[] R3S3 = { "2023,10,1,5", "2024,20,2,6", "2025,30,3,7" };
        private static readonly string[] R10 =
        {
            "2016,1", "2017,2", "2018,3", "2019,4", "2020,5",
            "2021,6", "2022,7", "2023,8", "2024,9", "2025,10"
        };

        private static string GenXml(
            string type,
            string title,
            string[] valueCols,
            string[] rows,
            string chartExtra = null,
            string colExtra = null,
            string dataFile = null)
        {
            var sb = new System.Text.StringBuilder();
            sb.Append("<ChartConfig><Chart><Type>").Append(type).Append("</Type>");
            if (title != null)
            {
                sb.Append("<Title>").Append(title).Append("</Title>");
            }

            if (chartExtra != null)
            {
                sb.Append(chartExtra);
            }

            sb.Append("<DataTable><Header><Column type=\"category\">年</Column>");
            foreach (string col in valueCols)
            {
                sb.Append("<Column type=\"value\"");
                if (colExtra != null)
                {
                    sb.Append(colExtra);
                }

                sb.Append(">").Append(col).Append("</Column>");
            }

            sb.Append("</Header>");
            if (dataFile != null)
            {
                sb.Append(dataFile);
            }
            else
            {
                sb.Append("<Rows>");
                foreach (string row in rows)
                {
                    sb.Append("<Row>").Append(row).Append("</Row>");
                }

                sb.Append("</Rows>");
            }

            sb.Append("</DataTable></Chart></ChartConfig>");
            return sb.ToString();
        }

        private static string CsvXml(string csvPath)
        {
            return GenXml(
                "column",
                TitleA,
                V1,
                null,
                null,
                null,
                "<DataFile fileName=\"" + csvPath + "\" hasHeader=\"true\"/>");
        }

        private static string DataXml2(string a, string b)
        {
            return GenXml("column", TitleA, V1, new[] { "2023," + a, "2024," + b });
        }

        private static string DataXml4(string a, string b, string c, string d)
        {
            return GenXml(
                "column",
                TitleA,
                V1,
                new[] { "2023," + a, "2024," + b, "2025," + c, "2026," + d });
        }

        private static string[] Rows(string y0)
        {
            return new[] { "2023," + y0, "2024,20", "2025,30" };
        }

        private static void AssertCreateSolo(TestRun run, WdSession s, string caseId, string xml, int? position = null)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, ConSent))
                {
                    run.CaseFail(caseId, "COM 写巩固句失败");
                    return;
                }

                s.ActivateDoc(path);
                s.GetContent(id);
                string err;
                Ok(run, caseId, CreateOk(s, id, xml, position, out err), err);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void AssertCreateSoloRaw(TestRun run, WdSession s, string caseId, string xml, object position)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, ConSent))
                {
                    run.CaseFail(caseId, "COM 写巩固句失败");
                    return;
                }

                s.ActivateDoc(path);
                s.GetContent(id);
                string err;
                Ok(run, caseId, CreateOkRaw(s, id, xml, position, out err), err);
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static void AssertCreateSoloFile(TestRun run, WdSession s, string caseId, string xmlPath)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, caseId))
            {
                return;
            }

            try
            {
                if (!s.TypePlain(path, ConSent))
                {
                    run.CaseFail(caseId, "COM 写巩固句失败");
                    return;
                }

                s.ActivateDoc(path);
                s.GetContent(id);
                ToolResult r = s.CreateChartFromXml(CreateArgs(id, xmlPath));
                Ok(run, caseId,
                    r.Success && WdSession.Str(r, "chart_id").StartsWith("C_", StringComparison.Ordinal),
                    r.Success ? "建图成功但无 C_" : (r.Error ?? "建图失败"));
            }
            finally
            {
                s.CloseDocByPath(path);
            }
        }

        private static bool CreateOk(WdSession s, string id, string xml, int? position, out string error)
        {
            var args = CreateArgs(id, WriteXml(xml));
            if (position.HasValue)
            {
                args["position"] = position.Value;
            }

            ToolResult r = s.CreateChartFromXml(args);
            error = r == null
                ? "空调用结果"
                : (r.Success
                    ? "建图成功但无 C_（message=" + WdSession.Str(r, "message") + "）"
                    : (r.Error ?? "未知错误"));
            return r != null
                && r.Success
                && WdSession.Str(r, "chart_id").StartsWith("C_", StringComparison.Ordinal);
        }

        private static bool CreateOkRaw(WdSession s, string id, string xml, object position, out string error)
        {
            var args = CreateArgs(id, WriteXml(xml));
            args["position"] = position;
            ToolResult r = s.CreateChartFromXml(args);
            error = r == null
                ? "空调用结果"
                : (r.Success
                    ? "建图成功但无 C_（message=" + WdSession.Str(r, "message") + "）"
                    : (r.Error ?? "未知错误"));
            return r != null
                && r.Success
                && WdSession.Str(r, "chart_id").StartsWith("C_", StringComparison.Ordinal);
        }

        private static Dictionary<string, object> ExtractArgsNoPath(string channelId, string chartId, string scope)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["scope"] = scope ?? "all",
                ["chart_selection"] = new Dictionary<string, object>
                {
                    ["type"] = "by_id",
                    ["chart_id"] = chartId ?? ""
                }
            };
        }

        private static Dictionary<string, object> ApplyArgs(
            string channelId,
            string chartId,
            string scope,
            string xmlContent,
            string mode = null)
        {
            var args = new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["chart_id"] = chartId ?? "",
                ["scope"] = scope,
                ["xml_content"] = xmlContent ?? ""
            };
            if (mode != null)
            {
                args["mode"] = mode;
            }

            return args;
        }

        private static Dictionary<string, object> DelArgs(string channelId, string chartId)
        {
            return new Dictionary<string, object>
            {
                ["channel_id"] = channelId,
                ["chart_id"] = chartId ?? ""
            };
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
            path = null;
            id = null;
            ToolResult opened = null;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                path = s.NewDocxPath();
                opened = s.Open(path, true, s.App);
                if (opened.Success)
                {
                    id = WdSession.Str(opened, "channel_id");
                    return true;
                }

                string err = opened.Error ?? "";
                bool permission = err.IndexOf("许可权", StringComparison.Ordinal) >= 0
                    || err.IndexOf("permission", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!permission || attempt == 2)
                {
                    FailMany(run, opened.Error ?? "新建失败", failIds);
                    return false;
                }

                WdSession.DismissOrphanExcel();
                System.Threading.Thread.Sleep(800);
            }

            FailMany(run, opened != null ? (opened.Error ?? "新建失败") : "新建失败", failIds);
            return false;
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

            args["path"] = string.IsNullOrEmpty(path)
                ? Path.Combine(Path.GetTempPath(), "ew-wh-x-" + Guid.NewGuid().ToString("N") + ".xml")
                : path;

            return args;
        }

        private static string ColumnXml(string title)
        {
            return TypedXml("column", title);
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
                "<ChartConfig><Chart><Type>column</Type><Title>" + TitleA + "</Title>"
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
                "<ChartConfig><Chart><Type>column</Type><Title>" + TitleA + "</Title>"
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
            return "<ChartConfig><Chart><Type>column</Type><Title>" + title + "</Title></Chart></ChartConfig>";
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
