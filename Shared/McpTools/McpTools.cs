using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// MCP内置工具管理器
    /// 提供各种内置工具的注册和执行功能
    /// </summary>
    public static class McpTools
    {
        /// <summary>
        /// 注册所有内置工具到指定的工具注册表
        /// </summary>
        /// <param name="toolRegistry">工具注册表</param>
        /// <param name="wordApplication">Word应用程序实例</param>
        public static void RegisterBuiltInTools(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication = null)
        {
            // 注册各个工具（metadata 由 Backend Registry 维护，此处仅注册执行 handler）
            // F_GetDocumentStructureTool 已下线：与 F_get_document_content（S_/T_ 快照）功能重叠，易误用
            // F_GetDocumentStructureTool.Register(toolRegistry, wordApplication);
            F_WriteFormatFileTool.Register(toolRegistry, wordApplication);
            F_ReadFileTool.Register(toolRegistry, wordApplication);
            F_CreateTableFromXmlTool.Register(toolRegistry, wordApplication);
            F_ExtractTableFormatTool.Register(toolRegistry, wordApplication);
            // F_ExtractTableDataTool / F_ApplyTableDataFromXmlTool 已下线：填空改走 F_read_table_row_values / F_apply_table_row_values
            F_ApplyTableFormatFromXmlTool.Register(toolRegistry, wordApplication);
            F_ReadTableRowValuesTool.Register(toolRegistry, wordApplication);
            F_ApplyTableRowValuesTool.Register(toolRegistry, wordApplication);
            F_OptimizeTableDisplayTool.Register(toolRegistry, wordApplication);
            F_DeleteTableTool.Register(toolRegistry, wordApplication);
            F_CreateChartFromXmlTool.Register(toolRegistry, wordApplication);
            F_ExtractChartXmlTool.Register(toolRegistry, wordApplication);
            F_ApplyChartFromXmlTool.Register(toolRegistry, wordApplication);
            F_DeleteChartTool.Register(toolRegistry, wordApplication);
            F_InsertSvgImageTool.Register(toolRegistry, wordApplication);
            F_InsertWorkspaceImageTool.Register(toolRegistry, wordApplication);
            F_ModifyYamlFileTool.Register(toolRegistry, wordApplication);
            F_CsvToXmlTool.Register(toolRegistry, wordApplication);
            F_GetDocumentContentTool.Register(toolRegistry, wordApplication);
            F_GetWorkbookContentTool.Register(toolRegistry, wordApplication);
            F_GetPresentationContentTool.Register(toolRegistry, wordApplication);
            F_ReadPptHtmlTool.Register(toolRegistry, wordApplication);
            F_ApplyPptHtmlTool.Register(toolRegistry, wordApplication);
            F_ManagePptSlideTool.Register(toolRegistry, wordApplication);
            F_PptAnimationTool.Register(toolRegistry, wordApplication);
            F_PptTransitionTool.Register(toolRegistry, wordApplication);
            F_ReadExcelRangeTool.Register(toolRegistry, wordApplication);
            F_WriteExcelRangeTool.Register(toolRegistry, wordApplication);
            F_ManageExcelSheetTool.Register(toolRegistry, wordApplication);
            F_ApplyExcelFormatTool.Register(toolRegistry, wordApplication);
            F_GetExcelFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyExcelConditionalFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyExcelStructureTool.Register(toolRegistry, wordApplication);
            F_ExcelPivotTool.Register(toolRegistry, wordApplication);
            F_ExcelChartTool.Register(toolRegistry, wordApplication);
            F_OpenDocumentTool.Register(toolRegistry, wordApplication);
            F_SwitchDefaultChannelTool.Register(toolRegistry, wordApplication);
            F_GetCurrDocChunkDisplayContentTool.Register(toolRegistry, wordApplication);
            F_GetCurrDocChunkParagraphDisplayContentTool.Register(toolRegistry, wordApplication);
            F_FormatTransferTool.Register(toolRegistry, wordApplication);
            F_ApplyPageSetupTool.Register(toolRegistry, wordApplication);
            F_ProcessDocumentActionsTool.Register(toolRegistry, wordApplication);
            // F_SetTrackRevisionsTool 已下线注册（2026-08）：实现保留，勿再 Register；WS 映射已移除
            // F_SetTrackRevisionsTool.Register(toolRegistry, wordApplication);
            F_ApplyDataProvenanceTool.Register(toolRegistry, wordApplication);
            F_GetDataProvenanceTool.Register(toolRegistry, wordApplication);
            F_InsertDocumentFromFileTool.Register(toolRegistry, wordApplication);
            F_ReplaceDocumentFromFileTool.Register(toolRegistry, wordApplication);
            F_GreetTool.Register(toolRegistry, wordApplication);
            // F_TestWordDocumentExtractor.Register(toolRegistry, wordApplication);
            F_TestParagraphCodeLocator.Register(toolRegistry, wordApplication);
            // F_TestHeadingRecognizer 已下线（开发测试用）
            // F_FindPositionTool 暂不下线注册：只读定位，Agent 暂不可用（代码保留）
            // F_FindPositionTool.Register(toolRegistry, wordApplication);
            F_CaptureDocumentPageImageTool.Register(toolRegistry, wordApplication);
            F_GetDocumentPageInfoTool.Register(toolRegistry, wordApplication);
            F_GetFormatContextTool.Register(toolRegistry, wordApplication);
            F_ApplyDocumentFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyParagraphFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyKbFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyKbParagraphFormatTool.Register(toolRegistry, wordApplication);
            F_ApplyKbTableFormatTool.Register(toolRegistry, wordApplication);
            F_ListDocumentOperationsTool.Register(toolRegistry, wordApplication);
            F_RestoreDocumentCheckpointTool.Register(toolRegistry, wordApplication);
        }
    }

    /// <summary>
    /// MCP工具定义（Word 端仅执行；LLM metadata 由 Backend Registry 维护）
    /// </summary>
    public class McpTool
    {
        public string name { get; set; }
        public Func<Dictionary<string, object>, Task<ToolResult>> Handler { get; set; }
    }

    /// <summary>
    /// 工具执行结果
    /// </summary>
    public class ToolResult
    {
        public bool Success { get; set; }
        public object Data { get; set; }
        public string Error { get; set; }
    }
}
