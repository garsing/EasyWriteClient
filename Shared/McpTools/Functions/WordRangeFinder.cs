using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Word Range查找器
    /// 用于在Word文档中查找指定文本的Range位置
    /// </summary>
    public static class WordRangeFinder
    {
        /// <summary>
        /// 将换行符转换为Word Find功能的特殊字符代码
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>转换后的文本（\r -> ^p, \n -> ^l, \t -> ^t）</returns>
        public static string ConvertNewlinesToWordCodes(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;

            // 先处理 \r\n 组合（转换为 ^p）
            string result = text.Replace("\r\n", "^p");
            
            // 再处理单独的 \r（段落标记）
            result = result.Replace("\r", "^p");
            
            // 最后处理单独的 \n（手动换行符）
            result = result.Replace("\n", "^l");

            // 制表符
            result = result.Replace("\t", "^t");
            
            return result;
        }

        /// <summary>
        /// 超长包围搜索：按 maxLength 切分，并去掉边界上被截断的 ^p 片段。
        /// 前缀末尾若为 ^p 或单独的 ^ 则删除；后缀开头若为 ^p 或单独的 p 则删除。
        /// </summary>
        private static void SplitForSurroundSearch(string convertedText, int maxLength, out string prefixSearchText, out string suffixSearchText)
        {
            prefixSearchText = TrimPrefixSurroundBoundary(convertedText.Substring(0, maxLength));
            suffixSearchText = TrimSuffixSurroundBoundary(convertedText.Substring(convertedText.Length - maxLength));
        }

        private static string TrimPrefixSurroundBoundary(string prefix)
        {
            if (string.IsNullOrEmpty(prefix))
            {
                return prefix;
            }

            if (prefix.Length >= 2 && prefix.EndsWith("^p", StringComparison.Ordinal))
            {
                return prefix.Substring(0, prefix.Length - 2);
            }

            if (prefix.EndsWith("^", StringComparison.Ordinal))
            {
                return prefix.Substring(0, prefix.Length - 1);
            }

            return prefix;
        }

        private static string TrimSuffixSurroundBoundary(string suffix)
        {
            if (string.IsNullOrEmpty(suffix))
            {
                return suffix;
            }

            if (suffix.StartsWith("^p", StringComparison.Ordinal))
            {
                return suffix.Substring(2);
            }

            if (suffix.StartsWith("p", StringComparison.Ordinal))
            {
                return suffix.Substring(1);
            }

            return suffix;
        }

        /// <summary>
        /// 为 Find 生成候选串：优先精确匹配；对旧映射（缺 \t）或 Word 多出的 Tab 尝试 ^t 变体。
        /// </summary>
        private static List<string> BuildFindTextCandidates(string convertedText)
        {
            var candidates = new List<string>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            void Add(string text)
            {
                if (!string.IsNullOrEmpty(text) && seen.Add(text))
                {
                    candidates.Add(text);
                }
            }

            Add(convertedText);

            if (convertedText.EndsWith("^p", StringComparison.Ordinal) && convertedText.Length > 2)
            {
                string withoutParagraph = convertedText.Substring(0, convertedText.Length - 2);
                if (!withoutParagraph.EndsWith("^t", StringComparison.Ordinal))
                {
                    Add(withoutParagraph + "^t^p");
                }
            }

            const int maxInnerSegmentLength = 30;
            for (int i = 0; i + 3 < convertedText.Length; i++)
            {
                if (convertedText[i] != '^' || convertedText[i + 1] != 'p')
                {
                    continue;
                }

                int innerStart = i + 2;
                int innerEnd = convertedText.IndexOf("^p", innerStart, StringComparison.Ordinal);
                if (innerEnd < 0 || innerEnd - innerStart > maxInnerSegmentLength)
                {
                    continue;
                }

                if (convertedText.IndexOf('^', innerStart, innerEnd - innerStart) >= 0)
                {
                    continue;
                }

                if (innerEnd >= 2
                    && convertedText[innerEnd - 2] == '^'
                    && convertedText[innerEnd - 1] == 't')
                {
                    continue;
                }

                Add(convertedText.Substring(0, innerEnd) + "^t^p" + convertedText.Substring(innerEnd + 2));
            }

            return candidates;
        }

        private static void ConfigureFind(Word.Find find, string searchText)
        {
            find.ClearFormatting();
            find.Text = searchText;
            find.MatchCase = true;
            find.MatchWholeWord = false;
            find.Wrap = Word.WdFindWrap.wdFindStop;
        }

        /// <summary>
        /// Find 命中长度与映射文本长度可能差 1～2 个不可见字符，将 Range 收束到期望长度。
        /// </summary>
        private static Word.Range AdjustRangeToExpectedLength(Word.Document doc, Word.Range found, string sentenceText)
        {
            if (doc == null || found == null || string.IsNullOrEmpty(sentenceText))
            {
                return found;
            }

            int expectedLength = sentenceText.Length;
            int actualLength = found.End - found.Start;
            if (actualLength == expectedLength)
            {
                return found;
            }

            const int invisibleCharTolerance = 3;
            if (actualLength > expectedLength && actualLength - expectedLength <= invisibleCharTolerance)
            {
                return doc.Range(found.Start, found.Start + expectedLength);
            }

            return found;
        }

        private static Word.Range FindFirstMatchInContent(
            Word.Document doc,
            List<string> candidates,
            string sentenceText,
            string debugInfo)
        {
            foreach (string candidate in candidates)
            {
                Word.Range searchRange = doc.Content;
                ConfigureFind(searchRange.Find, candidate);

                if (!searchRange.Find.Execute())
                {
                    continue;
                }

                if (candidate != candidates[0])
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}使用 Find 回退候选: \"{candidate}\"");
                }

                Word.Range sentenceRange = AdjustRangeToExpectedLength(doc, searchRange, sentenceText);
                System.Diagnostics.Debug.WriteLine(
                    $"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}✅ 找到匹配，Range位置: Start={sentenceRange.Start}, End={sentenceRange.End}");
                return sentenceRange;
            }

            return null;
        }

        // 暂时禁用：Word Find 在表内/跨 Story 时可能重复返回已匹配 Range，或在 searchStart 之后仍返回更早位置。
        // 该机制在合并单元格等场景下会导致 searchStart 逐字符推进、性能极差；暂容忍多处相同文本可能查不全。
        /*
        private static bool TryAdvanceSearchStartPastStaleFind(
            int foundStart,
            int foundEnd,
            int searchStart,
            int lastFoundStart,
            int lastFoundEnd,
            string logTag,
            out int nextSearchStart)
        {
            nextSearchStart = searchStart;
            bool duplicate = foundStart == lastFoundStart && foundEnd == lastFoundEnd;
            bool beforeWindow = foundStart < searchStart;
            if (!duplicate && !beforeWindow)
            {
                return false;
            }

            string reason = duplicate ? "重复位置" : "匹配在搜索窗口之前";
            int skip = Math.Max(foundEnd + 1, searchStart + 1);
            string prefix = string.IsNullOrEmpty(logTag) ? "" : $"({logTag}) ";
            if (skip <= searchStart)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindAllMatches] {prefix}⚠️ {reason} (Start={foundStart}, End={foundEnd})，无法推进，停止搜索");
                return true;
            }

            nextSearchStart = skip;
            System.Diagnostics.Debug.WriteLine(
                $"[FindAllMatches] {prefix}⚠️ {reason} (Start={foundStart}, End={foundEnd})，跳过继续 searchStart={nextSearchStart}");
            return true;
        }
        */

        /// <summary>
        /// 在整个文档中查找所有匹配指定文本的Range位置
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="searchText">搜索文本（已转换的Word代码格式）</param>
        /// <param name="debugInfo">调试信息</param>
        /// <returns>所有匹配的Range列表</returns>
        private static List<Word.Range> FindAllMatches(Word.Document doc, string searchText, string debugInfo = "")
        {
            foreach (string candidate in BuildFindTextCandidates(searchText))
            {
                List<Word.Range> matches = FindAllMatchesWithExactText(doc, candidate, debugInfo);
                if (matches.Count > 0)
                {
                    if (candidate != searchText)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}使用 Find 回退候选: \"{candidate}\"");
                    }

                    return matches;
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到 0 个匹配");
            return new List<Word.Range>();
        }

        private static List<Word.Range> FindAllMatchesWithExactText(Word.Document doc, string searchText, string debugInfo = "")
        {
            List<Word.Range> matches = new List<Word.Range>();
            
            try
            {
                int searchStart = 0;
                int maxIterations = 1000;
                int iterationCount = 0;
                
                while (searchStart < doc.Content.End && iterationCount < maxIterations)
                {
                    iterationCount++;
                    
                    Word.Range searchRange = doc.Range(searchStart, doc.Content.End);
                    ConfigureFind(searchRange.Find, searchText);
                    
                    if (searchRange.Find.Execute())
                    {
                        int foundStart = searchRange.Start;
                        int foundEnd = searchRange.End;

                        Word.Range foundRange = doc.Range(foundStart, foundEnd);
                        matches.Add(foundRange);

                        int newSearchStart = foundEnd + 1;
                        if (newSearchStart <= searchStart)
                        {
                            System.Diagnostics.Debug.WriteLine($"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}⚠️ 搜索位置未前进 (old={searchStart}, new={newSearchStart})，停止搜索");
                            break;
                        }

                        searchStart = newSearchStart;

                        System.Diagnostics.Debug.WriteLine($"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到匹配 {matches.Count}: Start={foundRange.Start}, End={foundRange.End}, nextSearchStart={searchStart}");
                    }
                    else
                    {
                        break;
                    }
                }
                
                if (iterationCount >= maxIterations)
                {
                    System.Diagnostics.Debug.WriteLine($"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}⚠️ 达到最大迭代次数 {maxIterations}，停止搜索");
                }
                
                if (matches.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到 {matches.Count} 个匹配");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FindAllMatches] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ 异常: {ex.GetType().Name} - {ex.Message}");
            }
            
            return matches;
        }

        /// <summary>
        /// 在指定 Range 内查找所有匹配（不超出 scope 边界）。
        /// </summary>
        private static List<Word.Range> FindAllMatchesInRange(
            Word.Document doc,
            Word.Range scope,
            string searchText,
            string debugInfo = "")
        {
            if (doc == null || scope == null || string.IsNullOrEmpty(searchText))
            {
                return new List<Word.Range>();
            }

            foreach (string candidate in BuildFindTextCandidates(searchText))
            {
                List<Word.Range> matches = FindAllMatchesInRangeWithExactText(doc, scope, candidate, debugInfo);
                if (matches.Count > 0)
                {
                    if (candidate != searchText)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[FindAllMatchesInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}使用 Find 回退候选: \"{candidate}\"");
                    }

                    return matches;
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[FindAllMatchesInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到 0 个匹配 scope={scope.Start}-{scope.End}");
            return new List<Word.Range>();
        }

        private static List<Word.Range> FindAllMatchesInRangeWithExactText(
            Word.Document doc,
            Word.Range scope,
            string searchText,
            string debugInfo = "")
        {
            List<Word.Range> matches = new List<Word.Range>();
            if (doc == null || scope == null || string.IsNullOrEmpty(searchText))
            {
                return matches;
            }

            try
            {
                int searchStart = scope.Start;
                int scopeEnd = scope.End;
                int maxIterations = 1000;
                int iterationCount = 0;

                while (searchStart < scopeEnd && iterationCount < maxIterations)
                {
                    iterationCount++;
                    Word.Range searchRange = doc.Range(searchStart, scopeEnd);
                    ConfigureFind(searchRange.Find, searchText);

                    if (searchRange.Find.Execute())
                    {
                        int foundStart = searchRange.Start;
                        int foundEnd = searchRange.End;

                        Word.Range foundRange = doc.Range(foundStart, foundEnd);
                        matches.Add(foundRange);

                        int newSearchStart = foundEnd + 1;
                        if (newSearchStart <= searchStart)
                        {
                            break;
                        }

                        searchStart = newSearchStart;
                    }
                    else
                    {
                        break;
                    }
                }

                if (matches.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[FindAllMatchesInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到 {matches.Count} 个匹配 scope={scope.Start}-{scopeEnd}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindAllMatchesInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ 异常: {ex.GetType().Name} - {ex.Message}");
            }

            return matches;
        }

        /// <summary>
        /// 长文本包围搜索：前 250 + 后 250 分别 Find，再遍历前缀/后缀 Range 组合并按长度校验。
        /// F_test_word_document_extractor（正文）、表内 P_ 定位、限定 Range 搜索均复用此函数。
        /// </summary>
        /// <param name="findMatches">在搜索域内查找子串，返回全部命中 Range（全文或限定 scope）</param>
        private static Word.Range FindLongTextRangeBySurroundSearch(
            Word.Document doc,
            string sentenceText,
            Func<string, string, List<Word.Range>> findMatches,
            string debugInfo = "",
            string logTag = "FindLongTextRangeBySurroundSearch")
        {
            if (doc == null || string.IsNullOrEmpty(sentenceText) || findMatches == null)
            {
                return null;
            }

            string convertedText = ConvertNewlinesToWordCodes(sentenceText);
            const int MAX_SEARCH_LENGTH = 250;

            System.Diagnostics.Debug.WriteLine(
                $"[{logTag}] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}原始文本长度: {sentenceText.Length}，转换后: {convertedText.Length}，使用包围搜索");

            SplitForSurroundSearch(convertedText, MAX_SEARCH_LENGTH, out string prefixSearchText, out string suffixSearchText);

            System.Diagnostics.Debug.WriteLine(
                $"[{logTag}] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}前250字符: \"{prefixSearchText}\"");
            System.Diagnostics.Debug.WriteLine(
                $"[{logTag}] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}后250字符: \"{suffixSearchText}\"");

            List<Word.Range> prefixRanges = findMatches(prefixSearchText, $"{debugInfo}_前缀");
            List<Word.Range> suffixRanges = findMatches(suffixSearchText, $"{debugInfo}_后缀");

            if (prefixRanges.Count == 0 || suffixRanges.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[{logTag}] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}未找到匹配：前缀{prefixRanges.Count}，后缀{suffixRanges.Count}");
                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[{logTag}] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}找到前缀 {prefixRanges.Count} 个，后缀 {suffixRanges.Count} 个");

            return MatchSurroundSearchCombinations(
                doc,
                prefixRanges,
                suffixRanges,
                sentenceText.Length,
                debugInfo);
        }

        /// <summary>
        /// 在指定 scope 内查找第一处完全落在 scope 内的命中（无 occurrenceIndex）。
        /// 短文本：BuildFindTextCandidates + ConfigureFind；窗外命中视为 stale 并跳过。
        /// 长文本：包围搜索，最终命中须 IsFullyInsideScope。
        /// </summary>
        public static Word.Range FindFirstInRange(Word.Range scope, string sentenceText, string debugInfo = "")
        {
            if (scope == null || string.IsNullOrEmpty(sentenceText))
            {
                return null;
            }

            Word.Document doc = scope.Document;
            if (doc == null)
            {
                return null;
            }

            string convertedText = ConvertNewlinesToWordCodes(sentenceText);
            const int MAX_SEARCH_LENGTH = 250;

            if (convertedText.Length <= MAX_SEARCH_LENGTH)
            {
                foreach (string candidate in BuildFindTextCandidates(convertedText))
                {
                    Word.Range hit = FindFirstMatchInScope(doc, scope, candidate, sentenceText, debugInfo);
                    if (hit != null)
                    {
                        if (candidate != convertedText)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}使用 Find 回退候选: \"{candidate}\"");
                        }

                        return hit;
                    }
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}未找到匹配 scope={scope.Start}-{scope.End}");
                return null;
            }

            Word.Range longHit = FindLongTextRangeBySurroundSearch(
                doc,
                sentenceText,
                (searchText, tag) => FindAllMatchesInRange(doc, scope, searchText, tag),
                debugInfo,
                "FindFirstInRange");

            if (longHit != null && IsFullyInsideScope(longHit, scope))
            {
                return longHit;
            }

            if (longHit != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}长文本命中落在 scope 外 Start={longHit.Start} End={longHit.End} scope={scope.Start}-{scope.End}");
            }

            return null;
        }

        private static Word.Range FindFirstMatchInScope(
            Word.Document doc,
            Word.Range scope,
            string candidate,
            string sentenceText,
            string debugInfo)
        {
            try
            {
                int scopeStart = scope.Start;
                int scopeEnd = scope.End;
                int searchStart = scopeStart;
                int maxIterations = 1000;
                int iterationCount = 0;

                while (searchStart < scopeEnd && iterationCount < maxIterations)
                {
                    iterationCount++;
                    Word.Range searchRange = doc.Range(searchStart, scopeEnd);
                    ConfigureFind(searchRange.Find, candidate);

                    if (!searchRange.Find.Execute())
                    {
                        break;
                    }

                    int foundStart = searchRange.Start;
                    int foundEnd = searchRange.End;
                    Word.Range foundRange = doc.Range(foundStart, foundEnd);

                    if (IsFullyInsideScope(foundRange, scope))
                    {
                        Word.Range adjusted = AdjustRangeToExpectedLength(doc, foundRange, sentenceText);
                        System.Diagnostics.Debug.WriteLine(
                            $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}✅ 找到匹配 Start={adjusted.Start} End={adjusted.End} scope={scopeStart}-{scopeEnd}");
                        return adjusted;
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}⚠️ 跳过 scope 外命中 Start={foundStart} End={foundEnd} scope={scopeStart}-{scopeEnd}");

                    int newSearchStart = foundEnd + 1;
                    if (newSearchStart <= searchStart)
                    {
                        break;
                    }

                    searchStart = newSearchStart;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindFirstInRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ 异常: {ex.GetType().Name} - {ex.Message}");
            }

            return null;
        }

        private static bool IsFullyInsideScope(Word.Range hit, Word.Range scope)
        {
            return hit != null && scope != null && hit.Start >= scope.Start && hit.End <= scope.End;
        }

        private static Word.Range FindSentenceRangeInScope(
            Word.Document doc,
            Word.Range scope,
            string sentenceText,
            string debugInfo = "")
        {
            return FindFirstInRange(scope, sentenceText, debugInfo);
        }

        /// <summary>
        /// 前缀/后缀包围搜索：遍历 Range 组合并按文本长度校验（与 FindSentenceRangeInDocument 一致）。
        /// </summary>
        private static Word.Range MatchSurroundSearchCombinations(
            Word.Document doc,
            List<Word.Range> prefixRanges,
            List<Word.Range> suffixRanges,
            int expectedLength,
            string debugInfo = "")
        {
            const int tolerance = 10;

            if (prefixRanges.Count == 1 && suffixRanges.Count == 1)
            {
                Word.Range prefixRange = prefixRanges[0];
                Word.Range suffixRange = suffixRanges[0];
                if (prefixRange.Start > suffixRange.End)
                {
                    return null;
                }

                Word.Range fullRange = doc.Range(prefixRange.Start, suffixRange.End);
                if (Math.Abs(fullRange.Text.Length - expectedLength) <= tolerance)
                {
                    return fullRange;
                }

                return null;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[MatchSurroundSearchCombinations] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}遍历组合：前缀{prefixRanges.Count} × 后缀{suffixRanges.Count}");

            int combinationIndex = 0;
            foreach (Word.Range prefixRange in prefixRanges)
            {
                foreach (Word.Range suffixRange in suffixRanges)
                {
                    combinationIndex++;
                    if (prefixRange.Start > suffixRange.End)
                    {
                        continue;
                    }

                    Word.Range fullRange = doc.Range(prefixRange.Start, suffixRange.End);
                    int actualLength = fullRange.Text.Length;
                    if (Math.Abs(actualLength - expectedLength) <= tolerance)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[MatchSurroundSearchCombinations] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}✅ 组合 {combinationIndex} Start={fullRange.Start} End={fullRange.End} 期望={expectedLength} 实际={actualLength}");
                        return fullRange;
                    }
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[MatchSurroundSearchCombinations] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}遍历 {combinationIndex} 个组合未匹配");
            return null;
        }

        /// <summary>
        /// 在整个文档中查找句子的第一处命中（无 occurrenceIndex）。
        /// </summary>
        public static Word.Range FindFirstInDocument(Word.Document doc, string sentenceText, string debugInfo = "")
        {
            return FindSentenceRangeInDocument(doc, sentenceText, debugInfo);
        }

        /// <summary>
        /// 在指定表格内查找句子的第一处命中（无 occurrenceIndex）。
        /// </summary>
        public static Word.Range FindFirstInTable(Word.Table table, string sentenceText, string debugInfo = "")
        {
            if (table == null || string.IsNullOrEmpty(sentenceText))
            {
                return null;
            }

            Word.Document doc = table.Range?.Document;
            if (doc == null)
            {
                return null;
            }

            List<Word.Range> cellRanges = CollectTableCellRanges(table, debugInfo);

            foreach (Word.Range cellRange in cellRanges)
            {
                Word.Range hit = FindFirstInRange(cellRange, sentenceText, debugInfo);
                if (hit != null)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[FindFirstInTable] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}Start={hit.Start} End={hit.End}");
                    return hit;
                }
            }

            if (cellRanges.Count == 0 && table.Range != null)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindFirstInTable] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}无可用Cell，回退 table.Range={table.Range.Start}-{table.Range.End}");
                return FindFirstInRange(table.Range, sentenceText, debugInfo + "_fallback");
            }

            System.Diagnostics.Debug.WriteLine(
                $"[FindFirstInTable] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}未找到匹配");
            return null;
        }

        /// <summary>
        /// 在整个文档中查找所有匹配指定文本的 Range（无 occurrenceIndex）。
        /// 短文本枚举全部命中；超长文本与 FindSentenceRangeInDocument 相同，至多返回一处。
        /// </summary>
        public static List<Word.Range> FindAllInDocument(Word.Document doc, string searchText, string debugInfo = "")
        {
            try
            {
                if (doc == null || string.IsNullOrEmpty(searchText))
                {
                    return new List<Word.Range>();
                }

                string convertedText = ConvertNewlinesToWordCodes(searchText);
                const int MAX_SEARCH_LENGTH = 250;

                if (convertedText.Length <= MAX_SEARCH_LENGTH)
                {
                    return FindAllMatches(doc, convertedText, debugInfo);
                }

                Word.Range single = FindLongTextRangeBySurroundSearch(
                    doc,
                    searchText,
                    (text, tag) => FindAllMatches(doc, text, tag),
                    debugInfo,
                    "FindAllInDocument");

                if (single == null)
                {
                    return new List<Word.Range>();
                }

                return new List<Word.Range> { single };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FindAllInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ 异常: {ex.GetType().Name} - {ex.Message}");
                return new List<Word.Range>();
            }
        }

        /// <summary>
        /// 在整个文档中查找句子的Range位置
        /// 支持超过250字符限制的文本：使用前250字符和后250字符进行包围搜索
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="sentenceText">句子文本</param>
        /// <param name="debugInfo">调试信息（可选，用于标识搜索的内容）</param>
        /// <returns>句子的Range，如果未找到则返回null</returns>
        public static Word.Range FindSentenceRangeInDocument(Word.Document doc, string sentenceText, string debugInfo = "")
        {
            try
            {
                if (doc == null || string.IsNullOrEmpty(sentenceText))
                {
                    System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}参数无效: doc={doc == null}, sentenceText为空={string.IsNullOrEmpty(sentenceText)}");
                    return null;
                }

                // 在函数开始处统一转换为Word代码格式，后续所有操作都基于转换后的文本
                string convertedText = ConvertNewlinesToWordCodes(sentenceText);
                int convertedLength = convertedText.Length;
                
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}原始文本长度: {sentenceText.Length}，转换后文本长度: {convertedLength}");

                // Word的Find功能对搜索字符串长度有限制（通常约255个字符）
                const int MAX_SEARCH_LENGTH = 250;
                
                // 如果转换后的文本长度不超过限制，使用标准搜索
                if (convertedLength <= MAX_SEARCH_LENGTH)
                {
                    System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}开始搜索，转换后文本: \"{convertedText}\"");
                    
                    List<string> candidates = BuildFindTextCandidates(convertedText);
                    Word.Range sentenceRange = FindFirstMatchInContent(doc, candidates, sentenceText, debugInfo);
                    if (sentenceRange != null)
                    {
                        return sentenceRange;
                    }

                    System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}未找到匹配");
                    return null;
                }
                else
                {
                    return FindLongTextRangeBySurroundSearch(
                        doc,
                        sentenceText,
                        (searchText, tag) => FindAllMatches(doc, searchText, tag),
                        debugInfo,
                        "FindSentenceRangeInDocument");
                }
            }
            catch (System.Runtime.InteropServices.COMException comEx)
            {
                string errorMsg = comEx.Message;
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ COM异常: {errorMsg}");
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 文本长度: {sentenceText?.Length ?? 0}");
                if (!string.IsNullOrEmpty(sentenceText))
                {
                    string escapedText = sentenceText.Replace("\r", "\\r").Replace("\n", "\\n");
                    System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 文本预览: {escapedText.Substring(0, Math.Min(100, escapedText.Length))}...");
                }
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}❌ 异常: {ex.GetType().Name} - {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 文本长度: {sentenceText?.Length ?? 0}");
                if (!string.IsNullOrEmpty(sentenceText))
                {
                    string escapedText = sentenceText.Replace("\r", "\\r").Replace("\n", "\\n");
                    System.Diagnostics.Debug.WriteLine($"[FindSentenceRangeInDocument] 文本预览: {escapedText.Substring(0, Math.Min(100, escapedText.Length))}...");
                }
                return null;
            }
        }

        /// <summary>
        /// 将 Find 命中 Range 上卷为整段 Paragraph.Range（0b I2）。
        /// </summary>
        public static Word.Range ExpandToParagraphRange(Word.Range hit)
        {
            if (hit == null)
            {
                return null;
            }

            try
            {
                Word.Paragraph para = hit.Paragraphs[1];
                return para?.Range?.Duplicate;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExpandToParagraphRange] 失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 收集表格内可访问的 Cell.Range。优先 table.Range.Cells（兼容纵向合并），回退 Rows 遍历。
        /// </summary>
        private static List<Word.Range> CollectTableCellRanges(Word.Table table, string debugInfo = "")
        {
            var cellRanges = new List<Word.Range>();
            if (table == null)
            {
                return cellRanges;
            }

            string tag = string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ";

            try
            {
                foreach (Word.Cell cell in table.Range.Cells)
                {
                    try
                    {
                        Word.Range cellRange = cell.Range;
                        if (cellRange != null && cellRange.Start < cellRange.End)
                        {
                            cellRanges.Add(cellRange);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[CollectTableCellRanges] {tag}跳过不可访问Cell（Range.Cells）");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[CollectTableCellRanges] {tag}Range.Cells 异常: {ex.Message}");
            }

            if (cellRanges.Count == 0)
            {
                try
                {
                    foreach (Word.Row row in table.Rows)
                    {
                        foreach (Word.Cell cell in row.Cells)
                        {
                            try
                            {
                                Word.Range cellRange = cell.Range;
                                if (cellRange != null && cellRange.Start < cellRange.End)
                                {
                                    cellRanges.Add(cellRange);
                                }
                            }
                            catch (System.Runtime.InteropServices.COMException)
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[CollectTableCellRanges] {tag}跳过不可访问Cell（Rows）");
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[CollectTableCellRanges] {tag}Rows 遍历异常: {ex.Message}");
                }
            }

            cellRanges.Sort((a, b) => a.Start.CompareTo(b.Start));
            System.Diagnostics.Debug.WriteLine(
                $"[CollectTableCellRanges] {tag}遍历到 {cellRanges.Count} 个可访问Cell");
            return cellRanges;
        }

        /// <summary>
        /// 移动光标到 Range 起点或终点，并滚入视口。
        /// </summary>
        public static bool MoveSelectionToRange(
            Word.Application app,
            Word.Range range,
            bool collapseToEnd = false,
            string debugInfo = "")
        {
            if (app == null || range == null)
            {
                return false;
            }

            try
            {
                int pos = collapseToEnd ? range.End : range.Start;
                if (app.Selection != null)
                {
                    app.Selection.SetRange(pos, pos);
                }

                Word.Document doc = range.Document;
                Word.Range scrollRange = doc != null ? doc.Range(pos, pos) : range;
                return ScrollRangeIntoView(app, scrollRange, debugInfo);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[MoveSelectionToRange] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 将 Range 滚入视口，不移动选区。
        /// </summary>
        public static bool ScrollRangeIntoView(Word.Application app, Word.Range range, string debugInfo = "")
        {
            if (app == null || range == null)
            {
                return false;
            }

            try
            {
                app.Activate();
                if (app.ActiveWindow != null)
                {
                    app.ActiveWindow.Activate();
                    object alignStart = true;
                    app.ActiveWindow.ScrollIntoView(range, alignStart);
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[ScrollRangeIntoView] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}Start={range.Start}, End={range.End}");
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ScrollRangeIntoView] {(string.IsNullOrEmpty(debugInfo) ? "" : $"({debugInfo}) ")}失败: {ex.Message}");
                return false;
            }
        }
    }
}

