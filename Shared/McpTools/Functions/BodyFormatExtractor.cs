using System;
using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 正文轻量格式抽取：unwrap 句池后边抽边判，写入 <see cref="DocumentState.BodyFormat"/>。
    /// 与全量 <see cref="FormatExtractor.ExtractSubtypeFormatMapping"/> 并列，互不覆盖。
    /// </summary>
    public static class BodyFormatExtractor
    {
        public const int MaxBodySamples = FormatExtractor.MaxBodySamplesForFormat;

        /// <summary>
        /// display→unwrap 句池→边抽边判→众数→DocumentState.BodyFormat。
        /// 失败只打日志，不抛给调用方。
        /// </summary>
        public static void ExtractAndSave(
            Word.Document doc,
            IList<string> chunkDisplayContents,
            Dictionary<string, string> nameToContent)
        {
            try
            {
                if (doc == null)
                {
                    DocumentState.SetBodyFormat(null);
                    if (FormatInheritHelper.EnableDbg)
                    {
                        System.Diagnostics.Debug.WriteLine("[BodyFormat] doc=null，清空 BodyFormat");
                    }
                    return;
                }

                List<SampleSentence> pool = SamplingSentencePool.BuildSamplingSentencePool(
                    chunkDisplayContents,
                    nameToContent);

                if (pool == null || pool.Count == 0)
                {
                    DocumentState.SetBodyFormat(null);
                    if (FormatInheritHelper.EnableDbg)
                    {
                        System.Diagnostics.Debug.WriteLine("[BodyFormat] pool 为空，清空 BodyFormat");
                    }
                    return;
                }

                var shuffled = pool.OrderBy(_ => Guid.NewGuid()).ToList();
                var samples = new List<Dictionary<string, object>>();
                var recognizer = new HeadingRecognizer();
                int tried = 0;
                int headingSkip = 0;
                int miss = 0;
                int index = 0;

                while (samples.Count < MaxBodySamples && index < shuffled.Count)
                {
                    SampleSentence item = shuffled[index++];
                    string text = item?.Text?.Trim() ?? "";
                    if (string.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    tried++;
                    Tuple<bool, string, string> heading = recognizer.IsHeading(text);
                    if (heading != null && heading.Item1)
                    {
                        headingSkip++;
                        continue;
                    }

                    Word.Range range = FindFirstSentenceMatch(doc, text);
                    if (range == null)
                    {
                        miss++;
                        continue;
                    }

                    Dictionary<string, object> format = FormatExtractor.ExtractFormatFromRange(range);
                    if (format != null && format.Count > 0)
                    {
                        samples.Add(format);
                    }
                }

                if (samples.Count == 0)
                {
                    DocumentState.SetBodyFormat(null);
                    if (FormatInheritHelper.EnableDbg)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[BodyFormat] pool_size={pool.Count} tried={tried} heading_skip={headingSkip} " +
                            $"miss={miss} samples=0 mode_keys= mode={{}}");
                        FormatInheritHelper.DbgLog("BodyFormatExtractor samples=0，BodyFormat 清空");
                    }
                    return;
                }

                Dictionary<string, object> mode = FormatExtractor.CalculateModeFormat(samples);
                DocumentState.SetBodyFormat(mode);

                // 自检日志规范（实现文档 §5.6）：[BodyFormat] pool/tried/… + mode 字段明细
                if (FormatInheritHelper.EnableDbg)
                {
                    string modeKeys = mode != null && mode.Count > 0
                        ? string.Join(",", mode.Keys)
                        : "";
                    string modeDebug = FormatInheritHelper.SnapshotToDebugString(mode);
                    System.Diagnostics.Debug.WriteLine(
                        $"[BodyFormat] pool_size={pool.Count} tried={tried} heading_skip={headingSkip} " +
                        $"miss={miss} samples={samples.Count} mode_keys={modeKeys} mode={modeDebug}");
                    FormatInheritHelper.DbgLog($"BodyFormatExtractor 众数已写入 DocumentState mode={modeDebug}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BodyFormat] ⚠️ 抽取异常（不阻断主流程）: {ex.Message}");
                try
                {
                    DocumentState.SetBodyFormat(null);
                }
                catch
                {
                    // ignore
                }
            }
        }

        /// <summary>I4：只取第一个命中 Range。</summary>
        private static Word.Range FindFirstSentenceMatch(Word.Document doc, string sentenceText)
        {
            try
            {
                if (doc == null || string.IsNullOrEmpty(sentenceText))
                {
                    return null;
                }

                string convertedText = WordRangeFinder.ConvertNewlinesToWordCodes(sentenceText);
                Word.Range searchRange = doc.Range(0, doc.Content.End);
                searchRange.Find.ClearFormatting();
                searchRange.Find.Text = convertedText;
                searchRange.Find.MatchCase = true;
                searchRange.Find.MatchWholeWord = false;
                searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;

                if (searchRange.Find.Execute())
                {
                    return doc.Range(searchRange.Start, searchRange.End);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BodyFormat] FindFirst 异常: {ex.Message}");
            }

            return null;
        }
    }
}
