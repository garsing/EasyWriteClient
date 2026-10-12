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
        public const int Target = 144;

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
                RunConsolidate(run, s);
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

        // ============ 巩固 100（wh-con-*）============

        private const string TitleC = "易写图再改";
        private const string ConSent = "易写图巩固句。";
        private const string PairSent = "易写图巩固乙句。";
        private const string TableCellText = "易写表巩固甲";

        private static readonly string[] ConNegIds =
        {
            "wh-con-n-root-case", "wh-con-n-no-chart", "wh-con-n-type-case", "wh-con-n-bar",
            "wh-con-n-area", "wh-con-n-pie", "wh-con-n-no-header", "wh-con-n-empty-header",
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
            if (s.IsWps)
            {
                RunConWpsRest(run, s);
                return;
            }

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
                    "<chartconfig><Chart><Type>column_clustered</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></chartconfig>"))),
                "ChartConfig");
            ExpectFail(run, "wh-con-n-no-chart",
                s.CreateChartFromXml(CreateArgs(null, WriteXml("<ChartConfig></ChartConfig>"))),
                "Chart");
            ExpectFail(run, "wh-con-n-type-case",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("LINE", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-bar",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("bar_clustered", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-area",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("area", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-pie",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(GenXml("pie", TitleA, V1, R3)))),
                "无效");
            ExpectFail(run, "wh-con-n-no-header",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Header");
            ExpectFail(run, "wh-con-n-empty-header",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable><Header></Header>"
                    + "<Rows><Row>2023,10</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "Column");
            ExpectFail(run, "wh-con-n-no-value",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column></Header>"
                    + "<Rows><Row>2023</Row></Rows></DataTable></Chart></ChartConfig>"))),
                "value");
            ExpectFail(run, "wh-con-n-no-rows",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "</DataTable></Chart></ChartConfig>"))),
                "Rows");
            ExpectFail(run, "wh-con-n-empty-rows",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable><Header>"
                    + "<Column type=\"category\">年</Column><Column type=\"value\">销量</Column></Header>"
                    + "<Rows></Rows></DataTable></Chart></ChartConfig>"))),
                "Row");
            ExpectFail(run, "wh-con-n-df-noattr",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type><DataTable><Header>"
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
                    GenXml("column_clustered", TitleA, V1, R3, null, " axisY=\"middle\"", null)))),
                "AxisY");
            ExpectFail(run, "wh-con-n-broken-xml",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    "<ChartConfig><Chart><Type>column_clustered</Type>"))),
                "XML");
            ExpectFail(run, "wh-con-n-format-only",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(FormatXml(TitleA)))),
                "SeriesCollection");
            ExpectFail(run, "wh-con-n-row-short",
                s.CreateChartFromXml(CreateArgs(null, WriteXml(
                    GenXml("column_clustered", TitleA, V1, new[] { "2023" })))),
                "Point");
        }

        /// <summary>建图特性 20 + position/杂项 9。一特性一稿，DataTable 行值互不相同。</summary>
        private static void RunConCreate(TestRun run, WdSession s)
        {
            AssertCreateSolo(run, s, "wh-con-x-no-title",
                GenXml("column_clustered", null, V1, Rows("101")));
            AssertCreateSolo(run, s, "wh-con-x-width",
                GenXml("column_clustered", TitleA, V1, Rows("102"), "<Width>500</Width>"));
            AssertCreateSolo(run, s, "wh-con-x-height",
                GenXml("column_clustered", TitleA, V1, Rows("103"), "<Height>300</Height>"));
            AssertCreateSolo(run, s, "wh-con-x-two-series",
                GenXml("column_clustered", TitleA, V2, new[] { "2023,10,1", "2024,20,2", "2025,30,3" }));
            AssertCreateSolo(run, s, "wh-con-x-three-series",
                GenXml("column_clustered", TitleA, V3, new[] { "2023,10,1,5", "2024,20,2,6", "2025,30,3,7" }));
            AssertCreateSolo(run, s, "wh-con-x-ten-rows",
                GenXml("column_clustered", TitleA, V1, R10));
            AssertCreateSolo(run, s, "wh-con-x-one-row",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,104" }));
            AssertCreateSolo(run, s, "wh-con-x-decimal",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,3.14", "2024,2.71", "2025,1.41" }));
            AssertCreateSolo(run, s, "wh-con-x-negative",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,-5", "2024,20", "2025,-30" }));
            AssertCreateSolo(run, s, "wh-con-x-big",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,1234567", "2024,20", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-empty-y",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,105", "2024,", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-nan-y",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,106", "2024,abc", "2025,30" }));
            AssertCreateSolo(run, s, "wh-con-x-fill-off",
                GenXml("column_clustered", TitleA, V1, new[] { "2023,107", "2024,", "2025,30" },
                    "<FillMissingValues>false</FillMissingValues>"));
            AssertCreateSolo(run, s, "wh-con-x-axisx",
                GenXml("column_clustered", TitleA, V1, Rows("108"), "<AxisX><Title>年份</Title></AxisX>"));
            AssertCreateSolo(run, s, "wh-con-x-axisy",
                GenXml("column_clustered", TitleA, V1, Rows("109"), "<AxisY><PrimaryTitle>万台</PrimaryTitle></AxisY>"));
            AssertCreateSolo(run, s, "wh-con-x-legend",
                GenXml("column_clustered", TitleA, V1, Rows("110"), "<LegendPosition>bottom</LegendPosition>"));
            AssertCreateSolo(run, s, "wh-con-x-datalabels",
                GenXml("column_clustered", TitleA, V1, Rows("111"), "<ShowDataLabels>true</ShowDataLabels>"));
            AssertCreateSolo(run, s, "wh-con-x-color",
                GenXml("column_clustered", TitleA, V1, Rows("112"), null, " color=\"#FF0000\"", null));

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
                        bool threeOk = CreateOk(s, id3, GenXml("column_clustered", TitleA, V1, Rows("201")), null, out err);
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
                GenXml("column_clustered", TitleA, V1, Rows("301")), 0);
            AssertCreateSolo(run, s, "wh-con-p-big",
                GenXml("column_clustered", TitleA, V1, Rows("302")), 99999);
            AssertCreateSoloRaw(run, s, "wh-con-p-nonint",
                GenXml("column_clustered", TitleA, V1, Rows("303")), "abc");

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
                        bool keepOk = CreateOk(s, idKeep, GenXml("column_clustered", TitleA, V1, Rows("304")), 0, out err);
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
                    + GenXml("column_clustered", TitleA, V1, Rows("305"))));

            string bomPath = Path.Combine(Path.GetTempPath(), "ew-wh-bom-" + Guid.NewGuid().ToString("N") + ".xml");
            File.WriteAllText(bomPath, GenXml("column_clustered", TitleA, V1, Rows("306")), new System.Text.UTF8Encoding(true));
            AssertCreateSoloFile(run, s, "wh-con-p-bom", bomPath);

            AssertCreateSolo(run, s, "wh-con-p-title-escape",
                GenXml("column_clustered", "易写图&amp;销量", V1, Rows("307")));
            AssertCreateSolo(run, s, "wh-con-p-title-cn",
                GenXml("column_clustered", "易写图「销量」（2025）", V1, Rows("308")));
            AssertCreateSolo(run, s, "wh-con-p-theme",
                GenXml("column_clustered", TitleA, V1, Rows("309"), "<Theme>Office</Theme>"));
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
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("401")))));
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
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("501")))));
            s.MoveSelectionToEnd(path);
            ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", "易写图巩固乙", V1, Rows("502")))));
            s.MoveSelectionToEnd(path);
            ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V2, new[] { "2023,10,1", "2024,20,2", "2025,30,3" }))));
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
            ToolResult c1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("601")))));
            s.MoveSelectionToEnd(path);
            ToolResult c2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("602")))));
            s.MoveSelectionToEnd(path);
            ToolResult c3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("603")))));
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
            ToolResult r1 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("611")))));
            s.MoveSelectionToEnd(path);
            ToolResult r2 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("612")))));
            s.MoveSelectionToEnd(path);
            ToolResult r3 = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("613")))));
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
            ToolResult fc = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("701")))));
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
            ToolResult ca = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("711")))));
            s.MoveSelectionToEnd(path);
            ToolResult cb = s.CreateChartFromXml(CreateArgs(id, WriteXml(GenXml("column_clustered", TitleA, V1, Rows("712")))));
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

        /// <summary>WPS：建/抽/套/删 unsupported 逐条断言；渠道/契约层负例两侧同一断言。</summary>
        private static void RunConWpsRest(TestRun run, WdSession s)
        {
            string path, id;
            if (!TryOpen(run, s, out path, out id, WpsConAll()))
            {
                return;
            }

            s.ActivateDoc(path);
            s.GetContent(id);

            ToolResult create = s.CreateChartFromXml(CreateArgs(id, WriteXml(ColumnXml(TitleA))));
            foreach (string cid in ConCreateIds)
            {
                ExpectFail(run, cid, create, "unsupported");
            }

            foreach (string cid in ConFlowIds)
            {
                ExpectFail(run, cid, create, "unsupported");
            }

            ToolResult extract = s.ExtractChartXml(ExtractArgs(id, "C_00001", null, "all"));
            foreach (string cid in ConExtractIds)
            {
                if (cid == "wh-con-e-unk-channel")
                {
                    continue;
                }

                ExpectFail(run, cid, extract, "unsupported");
            }

            ExpectFail(run, "wh-con-e-unk-channel",
                s.ExtractChartXml(ExtractArgs("w999", "C_00001", null, "all")),
                "未知");

            ToolResult apply = s.ApplyChartFromXml(ApplyArgs(id, "C_00001", "format", FormatXml(TitleB)));
            foreach (string cid in ConApplyIds)
            {
                if (cid == "wh-con-a-bad-xml" || cid == "wh-con-a-root"
                    || cid == "wh-con-a-unk-channel" || cid == "wh-con-a-empty-id")
                {
                    continue;
                }

                ExpectFail(run, cid, apply, "unsupported");
            }

            ExpectFail(run, "wh-con-a-bad-xml",
                s.ApplyChartFromXml(ApplyArgs(id, "C_00001", "format", "<ChartConfig><Chart>")),
                "XML");
            ExpectFail(run, "wh-con-a-root",
                s.ApplyChartFromXml(ApplyArgs(id, "C_00001", "format", "<NotChart></NotChart>")),
                "ChartConfig");
            ExpectFail(run, "wh-con-a-unk-channel",
                s.ApplyChartFromXml(ApplyArgs("w999", "C_00001", "format", FormatXml(TitleB))),
                "未知");
            ExpectFail(run, "wh-con-a-empty-id",
                s.ApplyChartFromXml(ApplyArgs(id, "", "format", FormatXml(TitleB))),
                "编号");

            ToolResult del = s.DeleteChart(DelArgs(id, "C_00001"));
            foreach (string cid in ConDeleteIds)
            {
                if (cid == "wh-con-d-unk-channel" || cid == "wh-con-d-empty-id")
                {
                    continue;
                }

                ExpectFail(run, cid, del, "unsupported");
            }

            ExpectFail(run, "wh-con-d-unk-channel", s.DeleteChart(DelArgs("w999", "C_00001")), "未知");
            ExpectFail(run, "wh-con-d-empty-id", s.DeleteChart(DelArgs(id, "")), "编号");

            s.CloseDocByPath(path);
        }

        private static string[] WpsConAll()
        {
            var all = new List<string>();
            all.AddRange(ConCreateIds);
            all.AddRange(ConExtractIds);
            all.AddRange(ConApplyIds);
            all.AddRange(ConDeleteIds);
            all.AddRange(ConFlowIds);
            return all.ToArray();
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
                "column_clustered",
                TitleA,
                V1,
                null,
                null,
                null,
                "<DataFile fileName=\"" + csvPath + "\" hasHeader=\"true\"/>");
        }

        private static string DataXml2(string a, string b)
        {
            return GenXml("column_clustered", TitleA, V1, new[] { "2023," + a, "2024," + b });
        }

        private static string DataXml4(string a, string b, string c, string d)
        {
            return GenXml(
                "column_clustered",
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

            args["path"] = string.IsNullOrEmpty(path)
                ? Path.Combine(Path.GetTempPath(), "ew-wh-x-" + Guid.NewGuid().ToString("N") + ".xml")
                : path;

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
