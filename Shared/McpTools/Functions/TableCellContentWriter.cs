using System;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public sealed class CellWriteResult
    {
        public bool Success { get; set; }
        public string Error { get; set; }
        public string Warning { get; set; }
    }

    public static class TableCellContentWriter
    {
        public static bool IsCellEmpty(Word.Cell cell)
        {
            if (cell == null)
            {
                return true;
            }

            try
            {
                return TableMergeHelper.IsNormalizedCellEmpty(cell.Range?.Text);
            }
            catch
            {
                return true;
            }
        }

        public static string NormalizePlainText(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "";
            }

            string noTags = Regex.Replace(text, "<[^>]+>", "");
            string decoded = WebUtility.HtmlDecode(noTags);
            return decoded
                .Replace("\r", "")
                .Replace("\a", "")
                .Trim();
        }

        public static string GetCellPlainText(Word.Cell cell)
        {
            if (cell == null)
            {
                return "";
            }

            try
            {
                return NormalizePlainText(cell.Range?.Text ?? "");
            }
            catch
            {
                return "";
            }
        }

        public static void ClearCellKeepOneParagraph(Word.Range cellRange)
        {
            if (cellRange == null)
            {
                return;
            }

            while (cellRange.Paragraphs.Count > 1)
            {
                cellRange.Paragraphs[cellRange.Paragraphs.Count].Range.Delete();
            }

            Word.Range firstPara = cellRange.Paragraphs[1].Range;
            firstPara.Text = "";
        }

        public static CellWriteResult WriteCell(Word.Cell cell, string content, string defaultFontColor = null)
        {
            var result = new CellWriteResult { Success = true };
            if (cell == null)
            {
                result.Success = false;
                result.Error = "cell 为空";
                return result;
            }

            string cellData = content ?? "";
            try
            {
                if (string.IsNullOrEmpty(NormalizePlainText(cellData)))
                {
                    ClearCellKeepOneParagraph(cell.Range);
                    return result;
                }

                if (cellData.Contains("<"))
                {
                    ClearCellKeepOneParagraph(cell.Range);
                    var htmlResult = InjectHtml(cell.Range, cellData);
                    if (!htmlResult.Success)
                    {
                        result.Success = false;
                        result.Error = htmlResult.Error;
                    }
                    else if (!string.IsNullOrEmpty(htmlResult.Warning))
                    {
                        result.Warning = htmlResult.Warning;
                    }

                    return result;
                }

                cell.Range.Text = cellData;
                if (!string.IsNullOrEmpty(defaultFontColor))
                {
                    try
                    {
                        cell.Range.Font.Color = (Word.WdColor)Convert.ToInt32(defaultFontColor, 16);
                    }
                    catch (Exception ex)
                    {
                        result.Warning = $"设置默认字体颜色失败: {ex.Message}";
                    }
                }

                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.Error = ex.Message;
                return result;
            }
        }

        public static CellWriteResult InjectHtml(Word.Range range, string htmlContent)
        {
            var result = new CellWriteResult { Success = true };
            try
            {
                var segments = HtmlParser.ParseHtml(htmlContent);
                bool success = HtmlParser.InsertHtmlSegments(range, segments);
                if (success)
                {
                    return result;
                }

                string plainText = Regex.Replace(htmlContent, "<[^>]+>", "");
                range.Text = plainText;
                result.Warning = "HTML解析失败，已降级为纯文本";
                return result;
            }
            catch (Exception ex)
            {
                try
                {
                    string plainText = Regex.Replace(htmlContent, "<[^>]+>", "");
                    range.Text = plainText;
                    result.Warning = $"HTML注入失败，已降级为纯文本: {ex.Message}";
                    return result;
                }
                catch (Exception ex2)
                {
                    result.Success = false;
                    result.Error = $"{ex.Message}；降级失败: {ex2.Message}";
                    return result;
                }
            }
        }
    }
}
