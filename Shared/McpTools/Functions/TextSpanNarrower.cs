using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 在已经落到的句 / 段 Range 里按字符串再收一截。不改文档里的表格结束符。
    /// </summary>
    public static class TextSpanNarrower
    {
        public static InsertTargetLocatorHelper.ResolveRangeResult TryNarrow(
            Word.Document doc,
            Word.Range range,
            TextSpan span,
            string code)
        {
            if (doc == null || range == null || span == null || string.IsNullOrEmpty(span.Value))
            {
                return Fail(TextSpanArg.Missing(code ?? "", span?.Value ?? ""));
            }

            int start = range.Start;
            int end = ShrinkOffCellEnd(doc, start, range.End);
            if (end < start)
            {
                end = start;
            }

            var hits = new List<Word.Range>();
            int cursor = start;
            int guard = 0;
            while (cursor < end && guard < 1000)
            {
                guard++;
                Word.Range window = doc.Range(cursor, end);
                Word.Range hit = WordRangeFinder.FindFirstInRange(window, span.Value, "text-span");
                if (hit == null || hit.Start < cursor || hit.End > end || hit.End <= cursor)
                {
                    break;
                }

                hits.Add(doc.Range(hit.Start, hit.End));
                cursor = hit.End;
            }

            int count = hits.Count;
            if (!span.HasIndex)
            {
                if (count == 0)
                {
                    return Fail(TextSpanArg.Missing(code, span.Value));
                }

                if (count >= 2)
                {
                    return Fail(TextSpanArg.Ambiguous(code, span.Value, count));
                }

                return Ok(hits[0]);
            }

            if (!span.IndexIsInteger || span.Index < 1 || span.Index > count)
            {
                return Fail(TextSpanArg.IndexOutOfRange(code, span.Value, count));
            }

            return Ok(hits[span.Index - 1]);
        }

        private static int ShrinkOffCellEnd(Word.Document doc, int start, int end)
        {
            if (end <= start + 1)
            {
                return end;
            }

            try
            {
                Word.Range probe = doc.Range(start, end);
                bool inTable = (bool)probe.Information[Word.WdInformation.wdWithInTable];
                if (!inTable)
                {
                    return end;
                }

                Word.Range cell = probe.Cells[1].Range;
                if (cell.End != end)
                {
                    return end;
                }

                return end - 1;
            }
            catch
            {
                return end;
            }
        }

        private static InsertTargetLocatorHelper.ResolveRangeResult Ok(Word.Range range)
        {
            return new InsertTargetLocatorHelper.ResolveRangeResult
            {
                Success = true,
                Range = range
            };
        }

        private static InsertTargetLocatorHelper.ResolveRangeResult Fail(string error)
        {
            return new InsertTargetLocatorHelper.ResolveRangeResult
            {
                Success = false,
                Error = error
            };
        }
    }
}
