using System;
using System.Collections.Generic;
using System.Text;

namespace WordAddIn1.SpreadsheetHost
{
    internal static class WorkbookPreviewMarkup
    {
        public const int MaxRows = 5;
        public const int MaxCols = 16;

        public static string BuildTable(
            string sheetName,
            string previewRange,
            bool truncated,
            IList<PreviewRow> rows)
        {
            var sb = new StringBuilder();
            sb.Append("<table id=\"");
            sb.Append(EscapeAttr((sheetName ?? "") + "!" + (previewRange ?? "")));
            sb.Append("\" sheet=\"");
            sb.Append(EscapeAttr(sheetName ?? ""));
            sb.Append("\" range=\"");
            sb.Append(EscapeAttr(previewRange ?? ""));
            sb.Append("\"");
            if (truncated)
            {
                sb.Append(" truncated=\"true\"");
            }

            sb.Append(">");
            if (rows != null)
            {
                foreach (PreviewRow row in rows)
                {
                    if (row == null)
                    {
                        continue;
                    }

                    sb.Append("<row r=\"");
                    sb.Append(row.RowNumber);
                    sb.Append("\">");
                    if (row.Cells != null)
                    {
                        foreach (PreviewCell cell in row.Cells)
                        {
                            if (cell == null)
                            {
                                continue;
                            }

                            sb.Append("<cell addr=\"");
                            sb.Append(EscapeAttr(cell.Addr ?? ""));
                            sb.Append("\"");
                            if (!string.IsNullOrEmpty(cell.Area) && cell.Area != cell.Addr)
                            {
                                sb.Append(" area=\"");
                                sb.Append(EscapeAttr(cell.Area));
                                sb.Append("\"");
                            }

                            if (cell.ColSpan > 1)
                            {
                                sb.Append(" colspan=\"");
                                sb.Append(cell.ColSpan);
                                sb.Append("\"");
                            }

                            if (cell.RowSpan > 1)
                            {
                                sb.Append(" rowspan=\"");
                                sb.Append(cell.RowSpan);
                                sb.Append("\"");
                            }

                            sb.Append(">");
                            sb.Append(EscapeText(cell.Text ?? ""));
                            sb.Append("</cell>");
                        }
                    }

                    sb.Append("</row>");
                }
            }

            sb.Append("</table>");
            return sb.ToString();
        }

        public static string BuildDisplayContents(WorkbookContentResult result)
        {
            var sb = new StringBuilder();
            sb.Append("工作簿：");
            sb.Append(result?.Name ?? "");
            sb.AppendLine();
            sb.Append("路径：");
            sb.Append(string.IsNullOrEmpty(result?.Path) ? "（未保存）" : result.Path);
            sb.AppendLine();
            sb.Append("kind=");
            sb.Append(result?.Kind ?? "");
            sb.Append(" channel_id=");
            sb.Append(result?.ChannelId ?? "");
            sb.AppendLine();
            if (result?.Sheets == null)
            {
                return sb.ToString();
            }

            foreach (WorkbookSheetInfo sheet in result.Sheets)
            {
                if (sheet == null)
                {
                    continue;
                }

                sb.AppendLine();
                sb.Append("## ");
                sb.Append(sheet.Name ?? "");
                if (string.Equals(sheet.SheetType, "chart", StringComparison.OrdinalIgnoreCase))
                {
                    sb.Append("  sheet_type=chart");
                    sb.AppendLine();
                    sb.AppendLine("（图表工作表，概览不读格子）");
                    continue;
                }

                if (sheet.Hidden)
                {
                    sb.Append("  hidden=true");
                }

                sb.Append("  used_range=");
                sb.Append(sheet.UsedRange ?? "");
                sb.Append(" last_row=");
                sb.Append(sheet.LastRow);
                sb.Append(" last_col=");
                sb.Append(sheet.LastCol ?? "");
                sb.AppendLine();
                if (!string.IsNullOrEmpty(sheet.Note))
                {
                    sb.AppendLine(sheet.Note);
                }

                sb.AppendLine(sheet.Preview ?? "");
            }

            return sb.ToString();
        }

        private static string EscapeAttr(string value)
        {
            return (value ?? "")
                .Replace("&", "&amp;")
                .Replace("\"", "&quot;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }

        private static string EscapeText(string value)
        {
            return (value ?? "")
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;");
        }
    }

    internal sealed class PreviewRow
    {
        public int RowNumber { get; set; }

        public List<PreviewCell> Cells { get; set; }
    }

    internal sealed class PreviewCell
    {
        public string Addr { get; set; }

        public string Area { get; set; }

        public int ColSpan { get; set; }

        public int RowSpan { get; set; }

        public string Text { get; set; }
    }
}
