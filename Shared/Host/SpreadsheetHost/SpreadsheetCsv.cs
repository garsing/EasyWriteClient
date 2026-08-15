using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordAddIn1.SpreadsheetHost
{
    /// <summary>
    /// 工作区 CSV 读写（与 F_csv_to_xml 同级引号规则；UTF-8 BOM 写出）。
    /// </summary>
    internal static class SpreadsheetCsv
    {
        public static List<List<string>> ParseFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return null;
            }

            // UTF-8（自动识别 BOM）
            string[] lines = File.ReadAllLines(filePath, Encoding.UTF8);
            var result = new List<List<string>>();
            foreach (string line in lines)
            {
                if (line == null)
                {
                    continue;
                }

                // 保留全空行对应的空字段行？写回需要对齐；跳过完全空行与 csv_to_xml 一致
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                result.Add(ParseLine(line));
            }

            return result;
        }

        public static void WriteFile(string filePath, IList<List<string>> rows)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("filePath");
            }

            string dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir))
            {
                Directory.CreateDirectory(dir);
            }

            var sb = new StringBuilder();
            if (rows != null)
            {
                for (int i = 0; i < rows.Count; i++)
                {
                    if (i > 0)
                    {
                        sb.Append("\r\n");
                    }

                    List<string> row = rows[i];
                    if (row == null)
                    {
                        continue;
                    }

                    for (int c = 0; c < row.Count; c++)
                    {
                        if (c > 0)
                        {
                            sb.Append(',');
                        }

                        sb.Append(EscapeField(row[c] ?? ""));
                    }
                }
            }

            // UTF-8 BOM，便于 Excel / 中文
            File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }

        /// <summary>
        /// 从读结果（稀疏 PreviewRow，合并延续格已跳过）铺成与 actualRange 同形的矩形网格。
        /// </summary>
        public static List<List<string>> BuildGridFromPreview(
            string actualRange,
            IList<PreviewRow> previewRows,
            out int rowCount,
            out int colCount,
            out string error)
        {
            rowCount = 0;
            colCount = 0;
            error = null;
            if (string.IsNullOrWhiteSpace(actualRange))
            {
                return new List<List<string>>();
            }

            if (!A1Address.TryParseRange(
                    actualRange.Trim(),
                    out int firstRow,
                    out int firstCol,
                    out int lastRow,
                    out int lastCol,
                    out error))
            {
                return null;
            }

            rowCount = lastRow - firstRow + 1;
            colCount = lastCol - firstCol + 1;
            var grid = new List<List<string>>(rowCount);
            for (int r = 0; r < rowCount; r++)
            {
                var line = new List<string>(colCount);
                for (int c = 0; c < colCount; c++)
                {
                    line.Add("");
                }

                grid.Add(line);
            }

            if (previewRows == null)
            {
                return grid;
            }

            foreach (PreviewRow row in previewRows)
            {
                if (row?.Cells == null)
                {
                    continue;
                }

                foreach (PreviewCell cell in row.Cells)
                {
                    if (cell == null || string.IsNullOrEmpty(cell.Addr))
                    {
                        continue;
                    }

                    if (!A1Address.TryParseCell(cell.Addr, out int sheetRow, out int sheetCol))
                    {
                        continue;
                    }

                    int rr = sheetRow - firstRow;
                    int cc = sheetCol - firstCol;
                    if (rr < 0 || cc < 0 || rr >= rowCount || cc >= colCount)
                    {
                        continue;
                    }

                    grid[rr][cc] = cell.Text ?? "";
                }
            }

            return grid;
        }

        public static List<string> ParseLine(string line)
        {
            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;
            char quoteChar = '"';

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];
                if (!inQuotes)
                {
                    if (c == '"' || c == '\'')
                    {
                        inQuotes = true;
                        quoteChar = c;
                    }
                    else if (c == ',')
                    {
                        result.Add(current.ToString());
                        current.Clear();
                    }
                    else
                    {
                        current.Append(c);
                    }
                }
                else if (c == quoteChar)
                {
                    if (i + 1 < line.Length && line[i + 1] == quoteChar)
                    {
                        current.Append(quoteChar);
                        i++;
                    }
                    else
                    {
                        inQuotes = false;
                    }
                }
                else
                {
                    current.Append(c);
                }
            }

            result.Add(current.ToString());
            return result;
        }

        private static string EscapeField(string value)
        {
            if (value == null)
            {
                value = "";
            }

            bool needQuotes = value.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0
                || value.StartsWith(" ", StringComparison.Ordinal)
                || value.EndsWith(" ", StringComparison.Ordinal);
            if (!needQuotes)
            {
                return value;
            }

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
