using System;
using System.Collections.Generic;
using System.Linq;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L6：readText chunk → S_/P_ 映射。
    /// S_ 与 P_ 并行切分（均不重不漏覆盖原文），再按偏移建立 S_↔P_。
    /// </summary>
    public static class MapReadTextToCodes
    {
        public static MapReadTextResult MapChunks(IReadOnlyList<string> chunks)
        {
            return MapChunks(chunks, new DocumentStateMappingStore());
        }

        public static MapReadTextResult MapChunks(IReadOnlyList<string> chunks, IMappingStore store)
        {
            store.BeginMappingPass();
            var mapped = new MapReadTextResult();

            foreach (string chunkWithLabel in chunks)
            {
                List<string> chunkParts = WordDocumentExtractor.SplitTextByCellsForMapping(
                    chunkWithLabel,
                    out List<int> chunkSegMarks);

                var chunkNonSegs = new List<string>();
                var chunkSegsSentenceNames = new List<List<string>>();
                var chunkSegsParagraphNames = new List<List<string>>();
                var chunkSegMarksList = new List<int>();
                var chunkAllSentenceNames = new List<string>();
                var chunkAllParagraphNames = new List<string>();

                int segMarkIdx = 0;
                for (int partIdx = 0; partIdx < chunkParts.Count; partIdx++)
                {
                    if (partIdx % 2 == 0)
                    {
                        chunkNonSegs.Add(chunkParts[partIdx]);
                        continue;
                    }

                    int segMark = segMarkIdx < chunkSegMarks.Count ? chunkSegMarks[segMarkIdx] : 0;
                    chunkSegMarksList.Add(segMark);
                    segMarkIdx++;

                    SegMapResult segResult = MapSegTextToCodes(chunkParts[partIdx], segMark, store);
                    chunkSegsSentenceNames.Add(segResult.SentenceNames);
                    chunkSegsParagraphNames.Add(segResult.ParagraphNames);
                    chunkAllSentenceNames.AddRange(segResult.SentenceNames);
                    chunkAllParagraphNames.AddRange(segResult.ParagraphNames);
                }

                var sentenceContentMap = ContentMapFromMappings(store.GetSentenceMappings());
                var paragraphContentMap = ContentMapFromMappings(store.GetParagraphMappings());

                mapped.AllChunkNonSegs.Add(chunkNonSegs);
                mapped.AllChunkSegMarks.Add(chunkSegMarksList);
                mapped.AllChunkSentenceNames.Add(chunkAllSentenceNames);
                mapped.AllChunkParagraphNames.Add(chunkAllParagraphNames);
                mapped.AllChunkSegsSentenceNames.Add(chunkSegsSentenceNames);
                mapped.AllChunkSegsParagraphNames.Add(chunkSegsParagraphNames);
                mapped.ChunkDisplayContents.Add(
                    BuildSentenceDisplayContent.Build(chunkNonSegs, chunkSegsSentenceNames, sentenceContentMap));
                mapped.ChunkParagraphDisplayContents.Add(
                    BuildParagraphDisplayContent.Build(chunkNonSegs, chunkSegsParagraphNames, paragraphContentMap));
            }

            mapped.SentenceMappings = CloneMappings(store.GetSentenceMappings());
            mapped.ParagraphMappings = CloneMappings(store.GetParagraphMappings());
            mapped.SentenceToParagraph = store.GetSentenceToParagraph()
                .ToDictionary(kv => kv.Key, kv => kv.Value);
            mapped.ParagraphToSentences = store.GetParagraphToSentences().ToDictionary(
                kv => kv.Key,
                kv => new List<string>(kv.Value));
            mapped.NameToSentenceContentMap = ContentMapFromMappings(store.GetSentenceMappings());
            return mapped;
        }

        public static SegMapResult MapSegTextToCodes(string segText, int segMark, IMappingStore store)
        {
            var result = new SegMapResult();
            if (store == null || string.IsNullOrEmpty(segText))
            {
                return result;
            }

            // Path P：按 \r 切段（StoredText 不重不漏覆盖原文）
            var paragraphRanges = new List<ParagraphRange>();
            int paragraphOffset = 0;
            foreach (ParagraphBlock block in SplitTextByParagraphs.Split(segText))
            {
                string stored = block.StoredText;
                var pPair = store.FindOrCreateParagraph(stored, segMark);
                string pName = pPair.Item1;
                result.ParagraphNames.Add(pName);
                paragraphRanges.Add(new ParagraphRange(pName, paragraphOffset, paragraphOffset + stored.Length));
                paragraphOffset += stored.Length;
            }

            // Path S：对完整 seg 并行分句（保留段末 \r，不走去 \r 后的 block.Text）
            int sentenceOffset = 0;
            foreach (string sentenceContent in SentenceSplitter.SplitTextBySentencesForDocument(segText))
            {
                if (sentenceContent == null)
                {
                    continue;
                }

                var sPair = store.FindOrCreateSentence(sentenceContent, segMark);
                string sName = sPair.Item1;
                result.SentenceNames.Add(sName);

                int sentenceStart = sentenceOffset;
                sentenceOffset += sentenceContent.Length;

                string paragraphName = FindParagraphNameAtOffset(paragraphRanges, sentenceStart);
                if (!string.IsNullOrEmpty(paragraphName))
                {
                    store.LinkSentenceToParagraph(sName, paragraphName);
                }
            }

            if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentMapping)
                && (paragraphOffset != segText.Length || sentenceOffset != segText.Length))
            {
                EasyWriteDiagnostics.Log(
                    DebugCategory.DocumentMapping,
                    $"不重不漏校验失败 segLen={segText.Length} pSum={paragraphOffset} sSum={sentenceOffset}");
            }

            return result;
        }

        public static SegMapResult MapSegTextToCodes(string segText, int segMark, InMemoryMappingStore store)
        {
            return MapSegTextToCodes(segText, segMark, (IMappingStore)store);
        }

        private static string FindParagraphNameAtOffset(
            IReadOnlyList<ParagraphRange> paragraphRanges,
            int sentenceStart)
        {
            if (paragraphRanges == null || paragraphRanges.Count == 0 || sentenceStart < 0)
            {
                return null;
            }

            for (int i = 0; i < paragraphRanges.Count; i++)
            {
                ParagraphRange range = paragraphRanges[i];
                if (sentenceStart >= range.Start && sentenceStart < range.End)
                {
                    return range.Name;
                }
            }

            // 句首恰好落在末尾边界时，归入最后一段
            ParagraphRange last = paragraphRanges[paragraphRanges.Count - 1];
            if (sentenceStart == last.End)
            {
                return last.Name;
            }

            return null;
        }

        private readonly struct ParagraphRange
        {
            public ParagraphRange(string name, int start, int end)
            {
                Name = name;
                Start = start;
                End = end;
            }

            public string Name { get; }
            public int Start { get; }
            public int End { get; }
        }

        private static List<List<string>> CloneMappings(IReadOnlyList<IReadOnlyList<string>> mappings)
        {
            var clone = new List<List<string>>();
            if (mappings == null)
            {
                return clone;
            }

            foreach (var row in mappings)
            {
                clone.Add(row != null ? new List<string>(row) : new List<string>());
            }

            return clone;
        }

        private static Dictionary<string, string> ContentMapFromMappings(IReadOnlyList<IReadOnlyList<string>> mappings)
        {
            var map = new Dictionary<string, string>();
            if (mappings == null)
            {
                return map;
            }

            foreach (var row in mappings)
            {
                if (row != null && row.Count >= 3)
                {
                    map[row[0]] = row[2];
                }
            }

            return map;
        }
    }
}
