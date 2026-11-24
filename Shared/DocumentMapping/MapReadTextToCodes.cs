using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// L6：readText chunk → S_/P_ 映射。
    /// S_ 与 P_ 并行切分（均不重不漏覆盖原文），再按偏移建立 S_↔P_。
    /// 这张对照表慎用，先保留，后续可能改。只按句子起点落段，处理不了句子黏连：
    /// 黏连会把后面的空段并进前一句，这句实际盖住多段，表里只记下起点那段，对应可能不准。
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
            LocateProbeLog.Line("MapChunks ResetPairing before " + TocFieldStamp.DumpInterior());
            TocFieldStamp.ResetPairing();
            var mapped = new MapReadTextResult();
            int chunkCount = chunks == null ? 0 : chunks.Count;
            var totalSw = Stopwatch.StartNew();

            var encodeSw = Stopwatch.StartNew();
            foreach (string chunkWithLabel in chunks)
            {
                List<string> chunkParts = WordDocumentExtractor.SplitTextByCellsForMapping(
                    chunkWithLabel,
                    out List<int> chunkSegMarks);

                var chunkNonSegs = new List<string>();
                var chunkSegsSentenceNames = new List<List<string>>();
                var chunkSegsParagraphNames = new List<List<string>>();
                var chunkSegsParagraphInners = new List<List<string>>();
                var chunkSegsSentenceFieldIds = new List<List<string>>();
                var chunkSegsParagraphFieldIds = new List<List<string>>();
                var chunkSegMarksList = new List<int>();
                var chunkAllSentenceNames = new List<string>();
                var chunkAllParagraphNames = new List<string>();

                int segMarkIdx = 0;
                string openTocId = null;
                for (int partIdx = 0; partIdx < chunkParts.Count; partIdx++)
                {
                    if (partIdx % 2 == 0)
                    {
                        openTocId = TocFieldStamp.ScanMarkers(chunkParts[partIdx], openTocId);
                        chunkNonSegs.Add(chunkParts[partIdx]);
                        continue;
                    }

                    int segMark = segMarkIdx < chunkSegMarks.Count ? chunkSegMarks[segMarkIdx] : 0;
                    chunkSegMarksList.Add(segMark);
                    segMarkIdx++;

                    SegMapResult segResult = MapSegTextToCodes(chunkParts[partIdx], segMark, store, openTocId);
                    chunkSegsSentenceNames.Add(segResult.SentenceNames);
                    chunkSegsParagraphNames.Add(segResult.ParagraphNames);
                    chunkSegsParagraphInners.Add(segResult.ParagraphInners);
                    chunkSegsSentenceFieldIds.Add(segResult.SentenceFieldIds);
                    chunkSegsParagraphFieldIds.Add(segResult.ParagraphFieldIds);
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
                mapped.AllChunkSegsSentenceFieldIds.Add(chunkSegsSentenceFieldIds);
                mapped.AllChunkSegsParagraphFieldIds.Add(chunkSegsParagraphFieldIds);
                IReadOnlyDictionary<string, string> paragraphDisplayMap =
                    DisplayMapFromMappings(store.GetParagraphMappings());
                mapped.ChunkParagraphDisplayContents.Add(
                    BuildParagraphDisplayContent.Build(
                        chunkNonSegs,
                        chunkSegsParagraphNames,
                        paragraphDisplayMap,
                        chunkSegsParagraphInners));
                mapped.ChunkParagraphAgentDisplays.Add(
                    BuildParagraphDisplayContent.BuildAgent(
                        chunkNonSegs,
                        chunkSegsParagraphNames,
                        paragraphDisplayMap,
                        chunkSegsParagraphInners));
            }

            mapped.NameToSentenceContentMap = ContentMapFromMappings(store.GetSentenceMappings());
            encodeSw.Stop();
            EasyWriteDiagnostics.LogTiming(
                "MapChunks.find_or_create",
                encodeSw.ElapsedMilliseconds,
                "chunks=" + chunkCount
                + " S_=" + mapped.NameToSentenceContentMap.Count);

            var uniqSw = Stopwatch.StartNew();
            StoryUniqEncode.Apply(mapped, store);
            uniqSw.Stop();

            mapped.SentenceMappings = CloneMappings(store.GetSentenceMappings());
            mapped.ParagraphMappings = CloneMappings(store.GetParagraphMappings());

            var displaySw = Stopwatch.StartNew();
            FinishSentenceDisplays(mapped);
            displaySw.Stop();
            EasyWriteDiagnostics.LogTiming(
                "MapChunks.finish_display",
                displaySw.ElapsedMilliseconds,
                "chunks=" + mapped.ChunkDisplayContents.Count);
            mapped.SentenceToParagraph = store.GetSentenceToParagraph()
                .ToDictionary(kv => kv.Key, kv => kv.Value);
            mapped.ParagraphToSentences = store.GetParagraphToSentences().ToDictionary(
                kv => kv.Key,
                kv => new List<string>(kv.Value));
            mapped.NameToSentenceContentMap = ContentMapFromMappings(store.GetSentenceMappings());
            FillParagraphSnapshots(mapped);
            totalSw.Stop();
            EasyWriteDiagnostics.LogTiming(
                "MapChunks.total",
                totalSw.ElapsedMilliseconds,
                "chunks=" + chunkCount
                + " find_or_create_ms=" + encodeSw.ElapsedMilliseconds
                + " uniq_ms=" + uniqSw.ElapsedMilliseconds
                + " display_ms=" + displaySw.ElapsedMilliseconds);
            LocateProbeLog.Line("MapChunks done " + TocFieldStamp.DumpInterior());
            return mapped;
        }

        private static void FillParagraphSnapshots(MapReadTextResult mapped)
        {
            var flat = new List<string>();
            var segParts = new List<string>();
            if (mapped.AllChunkSegsParagraphNames != null)
            {
                foreach (List<List<string>> chunkSegs in mapped.AllChunkSegsParagraphNames)
                {
                    List<List<string>> publicSegs = BuildParagraphDisplayContent.BuildPublicSegs(chunkSegs);
                    foreach (List<string> seg in publicSegs)
                    {
                        flat.AddRange(seg);
                        segParts.Add(string.Join(",", seg));
                    }
                }
            }

            mapped.ParagraphSnapshot = string.Join(",", flat);
            mapped.ParagraphSegSnapshot = string.Join(";", segParts);
        }

        private static void FinishSentenceDisplays(MapReadTextResult mapped)
        {
            mapped.ChunkDisplayContents.Clear();
            mapped.ChunkLocateDisplayContents.Clear();
            mapped.AllChunkSentenceNames.Clear();
            var content = DisplayMapFromMappings(mapped.SentenceMappings)
                ?? mapped.NameToSentenceContentMap;
            if (content == null || content.Count == 0)
            {
                content = mapped.NameToSentenceContentMap;
            }

            var inoperable = mapped.InoperableSentenceCodes
                ?? new HashSet<string>(StringComparer.Ordinal);
            var wrapSw = Stopwatch.StartNew();
            long wrapMs = 0;
            int wrapSets = 0;

            for (int c = 0; c < mapped.AllChunkNonSegs.Count; c++)
            {
                List<string> nonSegs = mapped.AllChunkNonSegs[c];
                List<List<string>> segs = c < mapped.AllChunkSegsSentenceNames.Count
                    ? mapped.AllChunkSegsSentenceNames[c]
                    : new List<List<string>>();
                string display = BuildSentenceDisplayContent.Build(nonSegs, segs, content, inoperable);
                List<List<string>> geoSets = new List<List<string>>();
                if (c < mapped.AllChunkSegNeedsGeometry.Count)
                {
                    List<bool> flags = mapped.AllChunkSegNeedsGeometry[c];
                    for (int s = 0; s < segs.Count; s++)
                    {
                        if (s < flags.Count && flags[s] && segs[s] != null && segs[s].Count > 0)
                        {
                            geoSets.Add(new List<string>(segs[s]));
                        }
                    }
                }

                if (geoSets.Count > 0)
                {
                    wrapSets += geoSets.Count;
                    wrapSw.Restart();
                    display = EmptyCellCode.WrapCellsMatchingCodeSets(display, geoSets);
                    wrapMs += wrapSw.ElapsedMilliseconds;
                    if (c < mapped.AllChunkSegNeedsGeometry.Count)
                    {
                        List<bool> flags = mapped.AllChunkSegNeedsGeometry[c];
                        for (int s = 0; s < segs.Count; s++)
                        {
                            if (s < flags.Count && flags[s])
                            {
                                segs[s] = new List<string>();
                            }
                        }
                    }
                }

                mapped.ChunkLocateDisplayContents.Add(display);
                mapped.ChunkDisplayContents.Add(TocFieldStamp.CollapseMarked(display));
                var flat = new List<string>();
                foreach (List<string> seg in segs)
                {
                    if (seg != null)
                    {
                        flat.AddRange(seg);
                    }
                }

                mapped.AllChunkSentenceNames.Add(flat);
            }

            if (wrapSets > 0)
            {
                EasyWriteDiagnostics.LogTiming(
                    "MapChunks.wrap_geometry",
                    wrapMs,
                    "sets=" + wrapSets);
            }
        }

        public static SegMapResult MapSegTextToCodes(string segText, int segMark, IMappingStore store, string tocId = null)
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
                string tagged = block.StoredText;
                string stored = AutonumConvert.StripTags(tagged);
                var pPair = store.FindOrCreateParagraph(stored, segMark);
                string pName = pPair.Item1;
                result.ParagraphInners.Add(tagged ?? "");
                RememberDisplayInner(store, pName, tagged);
                result.ParagraphNames.Add(pName);
                paragraphRanges.Add(new ParagraphRange(pName, paragraphOffset, paragraphOffset + stored.Length));
                paragraphOffset += stored.Length;
            }

            // Path S：对完整 seg 并行分句（保留段末 \r，不走去 \r 后的 block.Text）
            int sentenceOffset = 0;
            var sentenceParagraphIndex = new List<int>();
            foreach (string sentenceContent in SentenceSplitter.SplitTextBySentencesForDocument(segText))
            {
                if (sentenceContent == null)
                {
                    continue;
                }

                int sentencePara = FindParagraphIndexAtOffset(paragraphRanges, sentenceOffset);
                string storedSentence = AutonumConvert.StripTags(sentenceContent);
                var sPair = store.FindOrCreateSentence(storedSentence, segMark);
                string sName = sPair.Item1;
                RememberDisplayInner(store, sName, sentenceContent);
                result.SentenceNames.Add(sName);
                sentenceParagraphIndex.Add(sentencePara);
                sentenceOffset += storedSentence.Length;

                // 对照表慎用：只记句子起点所在段，黏连盖住的后面那些段不记。
                string paragraphName = sentencePara >= 0 ? paragraphRanges[sentencePara].Name : null;
                if (!string.IsNullOrEmpty(paragraphName))
                {
                    store.LinkSentenceToParagraph(sName, paragraphName);
                }
            }

            bool insideToc = !string.IsNullOrEmpty(tocId);
            for (int s = 0; s < result.SentenceNames.Count; s++)
            {
                result.SentenceFieldIds.Add(insideToc ? tocId : null);
                if (insideToc)
                {
                    TocFieldStamp.RememberInterior(result.SentenceNames[s]);
                }
            }

            for (int p = 0; p < result.ParagraphNames.Count; p++)
            {
                result.ParagraphFieldIds.Add(insideToc ? tocId : null);
                if (insideToc)
                {
                    TocFieldStamp.RememberInterior(result.ParagraphNames[p]);
                }
            }

            int storyLen = AutonumConvert.StripTags(segText).Length;
            if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentMapping)
                && (paragraphOffset != storyLen || sentenceOffset != storyLen))
            {
                EasyWriteDiagnostics.Log(
                    DebugCategory.DocumentMapping,
                    $"不重不漏校验失败 segLen={storyLen} pSum={paragraphOffset} sSum={sentenceOffset}");
            }

            return result;
        }

        public static SegMapResult MapSegTextToCodes(string segText, int segMark, InMemoryMappingStore store)
        {
            return MapSegTextToCodes(segText, segMark, (IMappingStore)store);
        }

        private static int FindParagraphIndexAtOffset(
            IReadOnlyList<ParagraphRange> paragraphRanges,
            int sentenceStart)
        {
            if (paragraphRanges == null || paragraphRanges.Count == 0 || sentenceStart < 0)
            {
                return -1;
            }

            for (int i = 0; i < paragraphRanges.Count; i++)
            {
                ParagraphRange range = paragraphRanges[i];
                if (sentenceStart >= range.Start && sentenceStart < range.End)
                {
                    return i;
                }
            }

            ParagraphRange last = paragraphRanges[paragraphRanges.Count - 1];
            if (sentenceStart == last.End)
            {
                return paragraphRanges.Count - 1;
            }

            return -1;
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

        private static void RememberDisplayInner(IMappingStore store, string name, string tagged)
        {
            store.SetDisplayInner(name, tagged ?? "");
        }

        private static Dictionary<string, string> DisplayMapFromMappings(IReadOnlyList<IReadOnlyList<string>> mappings)
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
                    map[row[0]] = row.Count >= 5 && !string.IsNullOrEmpty(row[4]) ? row[4] : row[2];
                }
            }

            return map;
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
