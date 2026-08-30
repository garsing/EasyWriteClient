using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 对已有表格 apply TableConfig XML，仅迁移 General（三字体 + Style，COM）。
    /// </summary>
    public static class TableConfigApplyHelper
    {
        public static TableConfigApplyResult ApplyTableConfigToExistingTable(
            Word.Document document,
            string tableId,
            string xmlContent,
            bool skipReadWord = false)
        {
            var result = new TableConfigApplyResult();
            if (document == null)
            {
                result.Success = false;
                result.Error = "文档不可用";
                return result;
            }
            if (string.IsNullOrWhiteSpace(tableId))
            {
                result.Success = false;
                result.Error = "未提供表格编号";
                return result;
            }
            if (string.IsNullOrWhiteSpace(xmlContent))
            {
                result.Success = false;
                result.Error = "未提供 XML 内容";
                return result;
            }

            if (!skipReadWord)
            {
                try
                {
                    WordReader.ReadWord(document);
                }
                catch (Exception ex)
                {
                    result.Success = false;
                    result.Error = $"生成表格映射表失败：{ex.Message}";
                    return result;
                }
            }

            Word.Table table = ResolveTableById(document, tableId, out string resolveError);
            if (table == null)
            {
                result.Success = false;
                result.Error = resolveError;
                return result;
            }

            var parseResult = TableConfigXmlParser.Parse(xmlContent);
            if (parseResult.Config == null)
            {
                result.Success = false;
                result.Error = parseResult.Error ?? "XML 解析失败";
                return result;
            }

            TableApplyConfig config = parseResult.Config;
            TableGeneralFormat general = config.General;

            if (general == null
                || (!TableGeneralFormatApplyHelper.HasAnyFontField(general)
                    && !TableGeneralFormatApplyHelper.ShouldApplyThreeLine(general.StyleConfig)))
            {
                result.Success = true;
                return result;
            }

            var applyResult = TableGeneralFormatApplyHelper.ApplyGeneralToTable(table, general);
            if (applyResult.Warnings != null && applyResult.Warnings.Count > 0)
            {
                result.Warnings.AddRange(applyResult.Warnings);
            }

            result.Success = applyResult.Success;
            result.Error = applyResult.Error;
            result.RowCount = applyResult.RowCount;
            result.ColumnCount = applyResult.ColumnCount;
            return result;
        }

        public static Word.Table ResolveTableById(Word.Document document, string tableId, out string error)
        {
            return TableResolveHelper.TryResolveTableById(document, tableId, out error);
        }
    }
}
