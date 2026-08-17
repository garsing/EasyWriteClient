using System;
using System.Collections.Generic;

namespace WordAddIn1
{
    /// <summary>
    /// 分类 Debug 类别常量（与后端 DEBUG_CATEGORIES 同名对齐）。
    /// </summary>
    internal static class DebugCategory
    {
        public const string Auth = "auth";
        public const string Ws = "ws";
        public const string Llm = "llm";
        public const string Orchestrator = "orchestrator";
        public const string Tool = "tool";
        public const string DocumentExtract = "document_extract";
        public const string DocumentMapping = "document_mapping";
        public const string Table = "table";
        public const string Chart = "chart";
        public const string Kb = "kb";
        public const string Workspace = "workspace";
        public const string Config = "config";
        public const string Misc = "misc";
        /// <summary>步骤耗时测评（Time / LogTiming）。</summary>
        public const string Timing = "timing";
        /// <summary>约定 HTML read/apply 诊断（形状撞号、几何漂移等）。</summary>
        public const string PptHtml = "ppt_html";
        /// <summary>WPS et 图表/透视 COM 排障旁路（chart_et_error.txt / pivot_et_error.txt）。</summary>
        public const string ExcelEt = "excel_et";
        /// <summary>知识库文档格式粗迁（F_format_transfer / FormatTransferHelper）卡点诊断。</summary>
        public const string FormatTransfer = "format_transfer";
        /// <summary>Desktop「打开文件」探测 / Word COM 事件。</summary>
        public const string OpenFiles = "open_files";

        public static readonly HashSet<string> Known = new HashSet<string>(StringComparer.Ordinal)
        {
            Auth,
            Ws,
            Llm,
            Orchestrator,
            Tool,
            DocumentExtract,
            DocumentMapping,
            Table,
            Chart,
            Kb,
            Workspace,
            Config,
            Misc,
            Timing,
            PptHtml,
            ExcelEt,
            FormatTransfer,
            OpenFiles
        };
    }
}
