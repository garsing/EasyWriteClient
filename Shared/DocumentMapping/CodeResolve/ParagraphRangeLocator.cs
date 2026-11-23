using System;
using System.Collections.Generic;
using WordAddIn1;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1.DocumentMapping.CodeResolve
{
    /// <summary>
    /// P_ 落到 Word 段的唯一入口：转字流 → 段号 → Paragraphs[k]，不 Find。
    /// </summary>
    public static class ParagraphRangeLocator
    {
        public static Word.Range TryResolve(
            Word.Document doc,
            string paragraphCode,
            string debugTag = null)
        {
            if (string.IsNullOrEmpty(paragraphCode))
            {
                return null;
            }

            return TryResolve(doc, new[] { paragraphCode }, null, debugTag);
        }

        public static Word.Range TryResolve(
            Word.Document doc,
            IReadOnlyList<string> codes,
            IReadOnlyList<string> domainOrderedCodes = null,
            string debugTag = null)
        {
            if (doc == null || codes == null || codes.Count == 0)
            {
                return null;
            }

            foreach (string code in codes)
            {
                if (!ParagraphPublicCode.TryParse(code, out _))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ParagraphRangeLocator] 公开码不认: {code}");
                    return null;
                }
            }

            IReadOnlyList<string> domain = domainOrderedCodes
                ?? ParagraphCodeAmbiguityHelper.GetOrderedParagraphCodesFromDisplay();
            if (domain == null || domain.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    "[ParagraphRangeLocator] paragraph display 为空");
                return null;
            }

            if (DisplaySequenceLocator.CountSubstringMatches(codes, domain) != 1)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphRangeLocator] display 不连续或非独一份 codes=[{string.Join(",", codes)}]");
                return null;
            }

            int firstK = 0;
            int lastK = 0;
            for (int i = 0; i < codes.Count; i++)
            {
                int k = DisplayParagraphIndexMap.ResolveParaIndex(codes[i]);
                if (k <= 0)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ParagraphRangeLocator] 未落到段号 P_={codes[i]}");
                    return null;
                }

                if (i == 0)
                {
                    firstK = k;
                }

                lastK = k;
            }

            string tag = string.IsNullOrEmpty(debugTag)
                ? string.Join(",", codes)
                : debugTag;

            try
            {
                int nPara = doc.Paragraphs.Count;
                if (firstK > nPara || lastK > nPara)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[ParagraphRangeLocator] 段号越界 {tag} → {firstK}-{lastK}/{nPara}");
                    return null;
                }

                int start = doc.Paragraphs[firstK].Range.Start;
                int end = doc.Paragraphs[lastK].Range.End;
                Word.Range range = ExcludeTrailingCellMarks(doc, doc.Range(start, end));
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphRangeLocator] {tag} → 段[{firstK}..{lastK}] Start={start} End={range.End}");
                return range;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ParagraphRangeLocator] Paragraphs 失败 {tag}: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 格内最后一段的 Range.End 会盖住格末/行末 \a。只削这个字符，正文段的 \r 不动。
        /// </summary>
        public static Word.Range ExcludeTrailingCellMarks(Word.Document doc, Word.Range range)
        {
            if (doc == null || range == null)
            {
                return range;
            }

            int start;
            int end;
            try
            {
                start = range.Start;
                end = range.End;
            }
            catch (Exception)
            {
                return range;
            }

            while (end > start)
            {
                string last;
                try
                {
                    last = doc.Range(end - 1, end).Text;
                }
                catch (Exception)
                {
                    break;
                }

                if (string.IsNullOrEmpty(last) || last[last.Length - 1] != '\a')
                {
                    break;
                }

                end--;
            }

            if (end == range.End)
            {
                return range;
            }

            try
            {
                return doc.Range(start, end);
            }
            catch (Exception)
            {
                return range;
            }
        }
    }
}
