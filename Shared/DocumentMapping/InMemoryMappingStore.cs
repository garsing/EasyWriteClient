using System;
using System.Collections.Generic;
using System.Security.Cryptography;
using System.Text;

namespace WordAddIn1.DocumentMapping
{
    /// <summary>
    /// 内存映射存储（单测 / 纯逻辑 map 用）。Hash 与 Python generate_content_hash（SHA256 hex）对齐。
    /// </summary>
    public sealed class InMemoryMappingStore : IMappingStore
    {
        public int NextSentenceIndex { get; set; }
        public int NextParagraphIndex { get; set; }
        public List<List<string>> SentenceMappings { get; } = new List<List<string>>();
        public List<List<string>> ParagraphMappings { get; } = new List<List<string>>();
        public Dictionary<string, string> SentenceToParagraph { get; } = new Dictionary<string, string>();
        public Dictionary<string, List<string>> ParagraphToSentences { get; } = new Dictionary<string, List<string>>();

        private readonly Dictionary<string, string> _sentenceHashToName = new Dictionary<string, string>();
        private readonly Dictionary<string, string> _paragraphHashToName = new Dictionary<string, string>();

        public void BeginMappingPass()
        {
            SentenceToParagraph.Clear();
            ParagraphToSentences.Clear();
        }

        public IReadOnlyList<IReadOnlyList<string>> GetSentenceMappings() => SentenceMappings;

        public IReadOnlyList<IReadOnlyList<string>> GetParagraphMappings() => ParagraphMappings;

        public IReadOnlyDictionary<string, string> GetSentenceToParagraph() => SentenceToParagraph;

        public IReadOnlyDictionary<string, List<string>> GetParagraphToSentences() => ParagraphToSentences;

        public static string GenerateContentHashHex(string content)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content ?? ""));
                var sb = new StringBuilder(hashBytes.Length * 2);
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }

                return sb.ToString();
            }
        }

        public Tuple<string, bool> FindOrCreateSentence(string content, int mark = 0)
        {
            string stored = content ?? "";
            string contentHash = GenerateContentHashHex(stored);
            if (_sentenceHashToName.TryGetValue(contentHash, out string existing))
            {
                UpdateMark(SentenceMappings, existing, mark);
                return Tuple.Create(existing, false);
            }

            string name = MappingCodeGenerator.GenerateSentenceName(NextSentenceIndex++);
            _sentenceHashToName[contentHash] = name;
            SentenceMappings.Add(new List<string> { name, contentHash, stored, mark.ToString() });
            return Tuple.Create(name, true);
        }

        public Tuple<string, bool> FindOrCreateParagraph(string storedText, int mark = 0)
        {
            string stored = storedText ?? "";
            string contentHash = GenerateContentHashHex(stored);
            if (_paragraphHashToName.TryGetValue(contentHash, out string existing))
            {
                UpdateMark(ParagraphMappings, existing, mark);
                return Tuple.Create(existing, false);
            }

            string name = MappingCodeGenerator.GenerateParagraphName(NextParagraphIndex++);
            _paragraphHashToName[contentHash] = name;
            ParagraphMappings.Add(new List<string> { name, contentHash, stored, mark.ToString() });
            return Tuple.Create(name, true);
        }

        public void LinkSentenceToParagraph(string sentenceName, string paragraphName)
        {
            SentenceToParagraph[sentenceName] = paragraphName;
            if (!ParagraphToSentences.TryGetValue(paragraphName, out List<string> bucket))
            {
                bucket = new List<string>();
                ParagraphToSentences[paragraphName] = bucket;
            }

            if (!bucket.Contains(sentenceName))
            {
                bucket.Add(sentenceName);
            }
        }

        private static void UpdateMark(List<List<string>> mappings, string name, int mark)
        {
            foreach (var row in mappings)
            {
                if (row.Count > 0 && row[0] == name)
                {
                    if (row.Count >= 4)
                    {
                        row[3] = mark.ToString();
                    }
                    else
                    {
                        row.Add(mark.ToString());
                    }

                    return;
                }
            }
        }
    }
}
