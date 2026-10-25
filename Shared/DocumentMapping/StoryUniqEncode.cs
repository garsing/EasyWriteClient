using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using WordAddIn1;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// 编码时在 DisplayStoryConvert.ToStory 上消歧：并段 FindOrCreate、完整格几何码、非格 D22。
    /// </summary>
    public static class StoryUniqEncode
    {
        public const string UnusableError = "ambiguous_unusable_sentence";
        public const string UnusableHint = "歧义无法操作";
        public const string UnusableStartTail = " 歧义无法操作";
        public const string FailBeforeWriteMessage = UnusableError + ": " + UnusableHint;

        private static readonly Regex D22Start = new Regex(
            @"\[(S_[0-9A-Za-z]+)-start 歧义无法操作\]",
            RegexOptions.Compiled);

        public static int CountOccurrences(string story, string needle)
        {
            if (string.IsNullOrEmpty(story) || string.IsNullOrEmpty(needle))
            {
                return 0;
            }

            int n = 0;
            int from = 0;
            while (from <= story.Length - needle.Length)
            {
                int i = story.IndexOf(needle, from, StringComparison.Ordinal);
                if (i < 0)
                {
                    break;
                }

                n++;
                from = i + 1;
            }

            return n;
        }

        public static void Apply(MapReadTextResult mapped, IMappingStore store)
        {
            if (mapped == null || store == null)
            {
                return;
            }

            mapped.InoperableSentenceCodes = new HashSet<string>(StringComparer.Ordinal);
            mapped.AllChunkSegNeedsGeometry = new List<List<bool>>();

            Dictionary<string, string> content = mapped.NameToSentenceContentMap;
            if (content == null || content.Count == 0)
            {
                content = ContentMapFromStore(store);
            }
            var toStorySw = Stopwatch.StartNew();
            string joined = string.Join("", mapped.ChunkParagraphDisplayContents ?? new List<string>());
            string story = DisplayStoryConvert.ToStory(joined);
            toStorySw.Stop();
            EasyWriteDiagnostics.LogTiming(
                "StoryUniq.to_story",
                toStorySw.ElapsedMilliseconds,
                "display_len=" + joined.Length + " story_len=" + (story == null ? 0 : story.Length));

            int segCount = 0;
            int sentenceIn = 0;
            int sentenceOut = 0;
            int geoCount = 0;
            var rewriteSw = Stopwatch.StartNew();
            for (int c = 0; c < mapped.AllChunkSegsSentenceNames.Count; c++)
            {
                List<List<string>> segs = mapped.AllChunkSegsSentenceNames[c];
                List<int> marks = c < mapped.AllChunkSegMarks.Count
                    ? mapped.AllChunkSegMarks[c]
                    : new List<int>();
                var geoFlags = new List<bool>();
                for (int s = 0; s < segs.Count; s++)
                {
                    int mark = s < marks.Count ? marks[s] : 0;
                    bool geometry;
                    int before = segs[s] == null ? 0 : segs[s].Count;
                    segs[s] = RewriteSeg(
                        segs[s],
                        mark,
                        story,
                        content,
                        store,
                        mapped.InoperableSentenceCodes,
                        out geometry);
                    segCount++;
                    sentenceIn += before;
                    sentenceOut += segs[s] == null ? 0 : segs[s].Count;
                    if (geometry)
                    {
                        geoCount++;
                    }

                    geoFlags.Add(geometry);
                }

                mapped.AllChunkSegNeedsGeometry.Add(geoFlags);
            }

            rewriteSw.Stop();
            mapped.NameToSentenceContentMap = ContentMapFromStore(store);
            DocumentState.SetInoperableSentenceCodes(mapped.InoperableSentenceCodes);
            EasyWriteDiagnostics.LogTiming(
                "StoryUniq.rewrite",
                rewriteSw.ElapsedMilliseconds,
                "segs=" + segCount
                + " sent_in=" + sentenceIn
                + " sent_out=" + sentenceOut
                + " geo=" + geoCount
                + " d22=" + mapped.InoperableSentenceCodes.Count);
        }

        public static string StartMarker(string sentenceName, bool inoperable)
        {
            if (inoperable)
            {
                return "[" + sentenceName + "-start" + UnusableStartTail + "]";
            }

            return "[" + sentenceName + "-start]";
        }

        public static bool CsvIsInoperable(string csv)
        {
            DocumentState.EnsureInoperableFromLastDisplay();
            return DocumentState.CsvContainsInoperable(csv);
        }

        public static void CollectInoperableFromDisplay(string display, HashSet<string> into)
        {
            if (string.IsNullOrEmpty(display) || into == null)
            {
                return;
            }

            foreach (Match m in D22Start.Matches(display))
            {
                string code = m.Groups[1].Value;
                if (!string.IsNullOrEmpty(code))
                {
                    into.Add(code);
                }
            }
        }

        private static List<string> RewriteSeg(
            List<string> originalNames,
            int segMark,
            string story,
            Dictionary<string, string> content,
            IMappingStore store,
            HashSet<string> inoperable,
            out bool needsGeometry)
        {
            needsGeometry = false;
            if (originalNames == null || originalNames.Count == 0)
            {
                return originalNames ?? new List<string>();
            }

            var stored = new List<string>(originalNames.Count);
            foreach (string name in originalNames)
            {
                content.TryGetValue(name, out string text);
                stored.Add(text ?? "");
            }

            var groups = new List<Group>();
            int i = 0;
            int n = stored.Count;
            while (i < n)
            {
                int j = i;
                string acc = stored[i];
                while (CountOccurrences(story, acc) != 1 && j + 1 < n)
                {
                    j++;
                    acc += stored[j];
                }

                groups.Add(new Group
                {
                    Start = i,
                    End = j,
                    Concat = acc,
                    Unique = CountOccurrences(story, acc) == 1
                });
                i = j + 1;
            }

            if (groups.Count >= 2 && !groups[groups.Count - 1].Unique)
            {
                Group last = groups[groups.Count - 1];
                Group prev = groups[groups.Count - 2];
                prev.End = last.End;
                prev.Concat += last.Concat;
                prev.Unique = true;
                groups.RemoveAt(groups.Count - 1);
            }

            if (groups.Count == 1 && !groups[0].Unique)
            {
                if (segMark == 1)
                {
                    needsGeometry = true;
                    return new List<string>(originalNames);
                }

                string code = store.FindOrCreateSentence(groups[0].Concat, segMark).Item1;
                inoperable.Add(code);
                content[code] = groups[0].Concat;
                return new List<string> { code };
            }

            var publicNames = new List<string>();
            foreach (Group g in groups)
            {
                if (g.Start == g.End && g.Unique)
                {
                    publicNames.Add(originalNames[g.Start]);
                    continue;
                }

                string merged = store.FindOrCreateSentence(g.Concat, segMark).Item1;
                content[merged] = g.Concat;
                publicNames.Add(merged);
            }

            return publicNames;
        }

        private static Dictionary<string, string> ContentMapFromStore(IMappingStore store)
        {
            var map = new Dictionary<string, string>();
            foreach (var row in store.GetSentenceMappings())
            {
                if (row != null && row.Count >= 3)
                {
                    map[row[0]] = row[2];
                }
            }

            return map;
        }

        private sealed class Group
        {
            public int Start;
            public int End;
            public string Concat;
            public bool Unique;
        }
    }
}
