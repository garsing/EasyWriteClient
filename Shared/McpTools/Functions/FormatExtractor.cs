using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 格式提取器：识别标题并提取格式信息
    /// </summary>
    public static class FormatExtractor
    {
        public const int MaxBodySamplesForFormat = 10;
        public const int MaxHeadingSamplesPerSubtype = 5;
        public const int MaxSamplesPerSubtypeInJson = 3;

        /// <summary>
        /// 将 detailed_subtype 格式映射写入 DocumentState（句子级标题映射清空），并写入按原文顺序的标题↔类型列表。
        /// </summary>
        /// <param name="subtypeFormatMapping">各 detailed_subtype 的格式（及 level 等）</param>
        /// <param name="classification">行级分类结果；为 null 时仅清空原文顺序标题列表</param>
        public static void SaveToDocumentState(
            List<Dictionary<string, object>> subtypeFormatMapping,
            DocumentLineClassificationResult classification = null)
        {
            DocumentState.SetSentenceNameToHeadingMap(new Dictionary<string, Dictionary<string, object>>());
            DocumentState.SetSubtypeFormatMapping(subtypeFormatMapping);
            List<Dictionary<string, object>> headingsOrdered = classification != null
                ? FlattenHeadingsInDocumentOrder(classification)
                : new List<Dictionary<string, object>>();
            DocumentState.SetRecognizedHeadingsInDocumentOrder(headingsOrdered);
            DebugPrintRecognizedHeadingsMapping(headingsOrdered);
        }

        /// <summary>
        /// 将已存入 DocumentState 的「原文顺序标题 ↔ 类型」映射打印到调试输出（可搜 <c>[FormatExtractor][标题映射]</c>）。
        /// </summary>
        private static void DebugPrintRecognizedHeadingsMapping(List<Dictionary<string, object>> headingsOrdered)
        {
            System.Diagnostics.Debug.WriteLine("");
            System.Diagnostics.Debug.WriteLine("==========================================");
            System.Diagnostics.Debug.WriteLine($"[FormatExtractor][标题映射] 已写入 DocumentState，共 {headingsOrdered?.Count ?? 0} 条（按原文行序）");
            System.Diagnostics.Debug.WriteLine("==========================================");
            if (headingsOrdered == null || headingsOrdered.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[FormatExtractor][标题映射] —— 无标题 ——");
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine("");
                return;
            }

            for (int i = 0; i < headingsOrdered.Count; i++)
            {
                Dictionary<string, object> h = headingsOrdered[i];
                string lineNum = h != null && h.ContainsKey("line") ? h["line"]?.ToString() ?? "?" : "?";
                string dst = h != null && h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                string sub = h != null && h.ContainsKey("subtype") ? h["subtype"]?.ToString() ?? "" : "";
                string lang = h != null && h.ContainsKey("language") ? h["language"]?.ToString() ?? "" : "";
                string lv = h != null && h.ContainsKey("level") ? h["level"]?.ToString() ?? "" : "";
                string rawText = h != null && h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                string preview = rawText.Replace("\r", "\\r").Replace("\n", "\\n");
                if (preview.Length > 120)
                {
                    preview = preview.Substring(0, 117) + "...";
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[FormatExtractor][标题映射]  #{i + 1,4}  line={lineNum,-5}  subtype={sub,-12}  detailed_subtype={dst,-36}  lang={lang,-10}  level={lv,-12}  text={preview}");
            }

            System.Diagnostics.Debug.WriteLine("[FormatExtractor][标题映射] —— 列表结束 ——");
            System.Diagnostics.Debug.WriteLine("==========================================");
            System.Diagnostics.Debug.WriteLine("");
        }

        /// <summary>
        /// 标题按 readText 切行后的行号（line）排序键；无法解析时靠后。
        /// 与 python_code/format_extractor.py 中 _heading_line_sort_key / flatten_headings_in_document_order 对齐。
        /// </summary>
        private static int GetHeadingLineSortKey(Dictionary<string, object> h)
        {
            if (h != null && h.TryGetValue("line", out object o) && o != null)
            {
                if (o is int ii)
                {
                    return ii;
                }

                if (int.TryParse(o.ToString(), out int p))
                {
                    return p;
                }
            }

            return int.MaxValue;
        }

        /// <summary>
        /// 将 HeadingsByLanguage 中所有标题合并，并按 line 升序排列（原文顺序）。
        /// </summary>
        private static List<Dictionary<string, object>> FlattenHeadingsInDocumentOrder(DocumentLineClassificationResult classification)
        {
            if (classification?.HeadingsByLanguage == null)
            {
                return new List<Dictionary<string, object>>();
            }

            List<Dictionary<string, object>> flat = classification.HeadingsByLanguage.SelectMany(kv => kv.Value).ToList();
            flat.Sort((a, b) => GetHeadingLineSortKey(a).CompareTo(GetHeadingLineSortKey(b)));
            return flat;
        }

        /// <summary>
        /// 按原文行序输出标题列表（不落库、不返回），仅 Debug 输出；可搜「按原文行序」。
        /// </summary>
        private static void ConsolePrintHeadingsInDocumentOrder(DocumentLineClassificationResult classification, string documentLabel)
        {
            List<Dictionary<string, object>> flat = FlattenHeadingsInDocumentOrder(classification);

            System.Diagnostics.Debug.WriteLine("");
            System.Diagnostics.Debug.WriteLine("==========================================");
            System.Diagnostics.Debug.WriteLine($"[FormatExtractor] 识别标题共 {flat.Count} 条（按原文行序），文档: {documentLabel ?? ""}");
            System.Diagnostics.Debug.WriteLine("==========================================");
            for (int i = 0; i < flat.Count; i++)
            {
                Dictionary<string, object> h = flat[i];
                string lineNum = h.ContainsKey("line") ? h["line"]?.ToString() ?? "?" : "?";
                string dst = h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                string lv = h.ContainsKey("level") ? h["level"]?.ToString() ?? "" : "";
                string lang = h.ContainsKey("language") ? h["language"]?.ToString() ?? "" : "";
                string rawText = h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                string preview = rawText.Replace("\r", "\\r").Replace("\n", "\\n");
                if (preview.Length > 120)
                {
                    preview = preview.Substring(0, 117) + "...";
                }

                System.Diagnostics.Debug.WriteLine(
                    $"  #{i + 1,4}  line={lineNum,-5}  level={lv,-12}  lang={lang,-10}  detailed_subtype={dst,-40}  text={preview}");
            }

            System.Diagnostics.Debug.WriteLine("[FormatExtractor] —— 标题列表结束 ——");
            System.Diagnostics.Debug.WriteLine("==========================================");
            System.Diagnostics.Debug.WriteLine("");
        }

        /// <summary>
        /// 调试输出：按 readText 按换行切分后的行号排序，打印当前识别到的全部标题
        /// </summary>
        private static void DebugWriteAllRecognizedHeadings(DocumentLineClassificationResult classification)
        {
            if (classification == null)
            {
                return;
            }

            List<Dictionary<string, object>> flat = FlattenHeadingsInDocumentOrder(classification);

            System.Diagnostics.Debug.WriteLine("[步骤3][全部标题] ==========================================");
                System.Diagnostics.Debug.WriteLine($"[步骤3][全部标题] 共 {flat.Count} 条（按 readText 按 \\r\\n|\\r|\\n 切行后的行号排序）");
            for (int i = 0; i < flat.Count; i++)
            {
                Dictionary<string, object> h = flat[i];
                string lineNum = h.ContainsKey("line") ? h["line"]?.ToString() ?? "?" : "?";
                string lang = h.ContainsKey("language") ? h["language"]?.ToString() ?? "" : "";
                string subtype = h.ContainsKey("subtype") ? h["subtype"]?.ToString() ?? "" : "";
                string dst = h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                string level = h.ContainsKey("level") ? h["level"]?.ToString() ?? "" : "";
                string rawText = h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                string textEsc = TextUtils.EscapeNewlinesForDb(rawText);
                System.Diagnostics.Debug.WriteLine(
                    $"[步骤3][全部标题] #{i + 1,4}  line={lineNum,-4}  lang={lang,-10}  subtype={subtype,-12}  detailed_subtype={dst,-32}  level={level,-10}  text={textEsc}");
            }

            System.Diagnostics.Debug.WriteLine("[步骤3][全部标题] ==========================================");
        }

        /// <summary>
        /// 基于 readText 或句级候选池抽样，在 Word 中提取格式并生成 subtype 映射。
        /// sentencePool 非 null 时走 pool 路径（含空列表）；为 null 时走 readText 行级分类。
        /// </summary>
        public static (DocumentLineClassificationResult classification, List<Dictionary<string, object>> subtypeFormatMapping) ExtractSubtypeFormatMapping(
            string readText,
            Word.Document doc,
            List<SampleSentence> sentencePool = null)
        {
            if (sentencePool != null)
            {
                return ExtractFromSentencePool(doc, sentencePool);
            }

            if (string.IsNullOrEmpty(readText) || doc == null)
            {
                var emptyClass = new DocumentLineClassificationResult(
                    new Dictionary<string, List<Dictionary<string, object>>>(),
                    new List<Dictionary<string, object>>());
                return (emptyClass, new List<Dictionary<string, object>>());
            }

            var step3StartTime = System.Diagnostics.Stopwatch.StartNew();
            var recognizer = new HeadingRecognizer();
            DocumentLineClassificationResult classification = recognizer.ClassifyLinesForFormatExtraction(readText);
            step3StartTime.Stop();

            int headingCount = classification.HeadingsByLanguage.Values.Sum(l => l.Count);
            System.Diagnostics.Debug.WriteLine($"[步骤3] ✅ 按行分类完成：标题 {headingCount} 行，正文候选行 {classification.BodyLines.Count} 行，耗时: {step3StartTime.ElapsedMilliseconds}ms");
            DebugWriteAllRecognizedHeadings(classification);
            ConsolePrintHeadingsInDocumentOrder(classification, string.IsNullOrEmpty(doc.Name) ? "未命名文档" : doc.Name);

            // 第五步：提取每个detailed_subtype的格式并计算众数
            System.Diagnostics.Debug.WriteLine("[步骤5] 提取每个detailed_subtype的格式并计算众数...");
            var step5StartTime = System.Diagnostics.Stopwatch.StartNew();
            
            // 5.0: 先提取正文格式（非标题行随机抽样）
            var step50StartTime = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Debug.WriteLine("[步骤5.0] 提取正文格式...");
            List<Dictionary<string, object>> bodyFormats = new List<Dictionary<string, object>>();
            
            if (classification.BodyLines.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[步骤5.0] 正文候选行 {classification.BodyLines.Count}，随机抽样至多 10 行");
                
                var randomBody = new Random();
                var sampledBodyLines = classification.BodyLines.OrderBy(_ => randomBody.Next()).Take(Math.Min(MaxBodySamplesForFormat, classification.BodyLines.Count)).ToList();
                
                int bodySampleSkippedEmpty = 0;
                int bodySampleSearched = 0;
                int bodySampleLinesHit = 0;
                int bodySampleLinesMiss = 0;

                foreach (var bodyInfo in sampledBodyLines)
                {
                    string sentenceText = bodyInfo.ContainsKey("text") ? bodyInfo["text"]?.ToString() ?? "" : "";
                    if (string.IsNullOrEmpty(sentenceText))
                    {
                        bodySampleSkippedEmpty++;
                        continue;
                    }
                    
                    bodySampleSearched++;
                    List<Word.Range> foundRanges = FindAllSentenceMatches(doc, sentenceText);
                    int n = foundRanges.Count;
                    if (n > 0)
                    {
                        bodySampleLinesHit++;
                    }
                    else
                    {
                        bodySampleLinesMiss++;
                    }
                    
                    foreach (Word.Range range in foundRanges)
                    {
                        Dictionary<string, object> format = ExtractFormatFromRange(range);
                        bodyFormats.Add(format);
                    }
                }
                
                System.Diagnostics.Debug.WriteLine(
                    $"[步骤5.0]   正文抽样 {sampledBodyLines.Count} 行：空文本跳过 {bodySampleSkippedEmpty}，实际参与查找 {bodySampleSearched} 行；" +
                    $"其中至少命中 1 处 {bodySampleLinesHit} 行，未命中 {bodySampleLinesMiss} 行；文档命中合计 {bodyFormats.Count} 个 Range（即格式样本数）");
            }
            
            // 计算正文格式的众数
            Dictionary<string, object> bodyModeFormat = bodyFormats.Count > 0 ? CalculateModeFormat(bodyFormats) : new Dictionary<string, object>();
            step50StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5.0] ✅ 正文格式提取完成，耗时: {step50StartTime.ElapsedMilliseconds}ms");
            
            // 5.1: 为所有可能的标题类型生成默认格式
            var step51StartTime = System.Diagnostics.Stopwatch.StartNew();
            System.Diagnostics.Debug.WriteLine("[步骤5.1] 为所有可能的标题类型生成默认格式...");
            Dictionary<string, Dictionary<string, object>> defaultFormatsBySubtype = new Dictionary<string, Dictionary<string, object>>();
            List<string> allPossibleSubtypes = GetAllPossibleDetailedSubtypes();
            
            foreach (string detailedSubtype in allPossibleSubtypes)
            {
                Dictionary<string, object> defaultFormat = GenerateDefaultFormat(bodyModeFormat, detailedSubtype);
                defaultFormatsBySubtype[detailedSubtype] = defaultFormat;
            }
            step51StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5.1] ✅ 为 {allPossibleSubtypes.Count} 个标题类型生成了默认格式，耗时: {step51StartTime.ElapsedMilliseconds}ms");
            
            // 5.2: 按 detailed_subtype 分组（行级标题，合成键仅用于抽样）
            var step52StartTime = System.Diagnostics.Stopwatch.StartNew();
            Dictionary<string, List<KeyValuePair<string, Dictionary<string, object>>>> sentencesByDetailedSubtype = new Dictionary<string, List<KeyValuePair<string, Dictionary<string, object>>>>();
            
            int headingOrdinal = 0;
            foreach (var langList in classification.HeadingsByLanguage.Values)
            {
                foreach (Dictionary<string, object> info in langList)
                {
                    string detailedSubtype = info.ContainsKey("detailed_subtype") ? info["detailed_subtype"]?.ToString() ?? "" : "";
                    if (string.IsNullOrEmpty(detailedSubtype))
                    {
                        continue;
                    }

                    if (!sentencesByDetailedSubtype.ContainsKey(detailedSubtype))
                    {
                        sentencesByDetailedSubtype[detailedSubtype] = new List<KeyValuePair<string, Dictionary<string, object>>>();
                    }

                    var headingWithType = new Dictionary<string, object>(info);
                    headingWithType["type"] = "heading";
                    string syntheticKey = "H_" + headingOrdinal++;
                    sentencesByDetailedSubtype[detailedSubtype].Add(new KeyValuePair<string, Dictionary<string, object>>(syntheticKey, headingWithType));
                }
            }
            step52StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5.2] ✅ 找到 {sentencesByDetailedSubtype.Count} 个有标题行对应的 detailed_subtype，耗时: {step52StartTime.ElapsedMilliseconds}ms");
            
            // 5.3: 提取格式（标题类型抽样5个）
            var step53StartTime = System.Diagnostics.Stopwatch.StartNew();
            Dictionary<string, List<Dictionary<string, object>>> formatListBySubtype = new Dictionary<string, List<Dictionary<string, object>>>();
            
            int totalSentencesProcessed = 0;
            int totalRangesFound = 0;
            int totalFormatsExtracted = 0;
            
            // 遍历每个detailed_subtype（标题类型）
            foreach (var kvp in sentencesByDetailedSubtype)
            {
                string detailedSubtype = kvp.Key;
                List<KeyValuePair<string, Dictionary<string, object>>> sentencesWithName = kvp.Value;
                
                var subtypeStartTime = System.Diagnostics.Stopwatch.StartNew();
                System.Diagnostics.Debug.WriteLine($"[步骤5.3] 处理 detailed_subtype: {detailedSubtype}，包含 {sentencesWithName.Count} 个标题行，随机抽样至多 5 个");
                
                List<KeyValuePair<string, Dictionary<string, object>>> sentencesToSample = sentencesWithName;
                
                System.Diagnostics.Debug.WriteLine($"[步骤5.3]   待抽样 {sentencesToSample.Count} 个标题行");
                
                List<Dictionary<string, object>> formatsForThisSubtype = new List<Dictionary<string, object>>();
                
                // 随机抽样5个句子
                var random = new Random();
                var sampledSentences = sentencesToSample.OrderBy(x => random.Next()).Take(Math.Min(MaxHeadingSamplesPerSubtype, sentencesToSample.Count)).ToList();
                
                // 遍历抽样后的句子
                foreach (var item in sampledSentences)
                {
                    var sentenceInfo = item.Value;
                    string sentenceText = sentenceInfo.ContainsKey("text") ? sentenceInfo["text"]?.ToString() ?? "" : "";
                    if (string.IsNullOrEmpty(sentenceText))
                    {
                        continue;
                    }
                    
                    totalSentencesProcessed++;
                    
                    // 在Word中搜索该句子（可能找到多个匹配）
                    var searchStartTime = System.Diagnostics.Stopwatch.StartNew();
                    List<Word.Range> foundRanges = FindAllSentenceMatches(doc, sentenceText);
                    searchStartTime.Stop();
                    
                    totalRangesFound += foundRanges.Count;
                    
                    if (foundRanges.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[步骤5.3]   句子 \"{sentenceText.Substring(0, Math.Min(30, sentenceText.Length))}...\" 找到 {foundRanges.Count} 个匹配，搜索耗时: {searchStartTime.ElapsedMilliseconds}ms");
                    }
                    
                    // 对每个找到的Range提取格式
                    var extractStartTime = System.Diagnostics.Stopwatch.StartNew();
                    foreach (Word.Range range in foundRanges)
                    {
                        Dictionary<string, object> format = ExtractFormatFromRange(range);
                        formatsForThisSubtype.Add(format);
                        totalFormatsExtracted++;
                    }
                    extractStartTime.Stop();
                    
                    if (foundRanges.Count > 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"[步骤5.3]     提取格式耗时: {extractStartTime.ElapsedMilliseconds}ms（平均每个Range: {(foundRanges.Count > 0 ? extractStartTime.ElapsedMilliseconds / (double)foundRanges.Count : 0):F2}ms）");
                    }
                }
                
                subtypeStartTime.Stop();
                
                if (formatsForThisSubtype.Count > 0)
                {
                    formatListBySubtype[detailedSubtype] = formatsForThisSubtype;
                    System.Diagnostics.Debug.WriteLine($"[步骤5.3]   detailed_subtype {detailedSubtype} 共收集到 {formatsForThisSubtype.Count} 个格式样本，总耗时: {subtypeStartTime.ElapsedMilliseconds}ms");
                }
            }
            step53StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5.3] ✅ 格式提取完成，处理了 {totalSentencesProcessed} 个标题句子（抽样），找到 {totalRangesFound} 个Range，提取了 {totalFormatsExtracted} 个格式样本，总耗时: {step53StartTime.ElapsedMilliseconds}ms");
            
            // 5.4: 计算每个detailed_subtype的众数格式（排除与正文格式一致的格式），如果存在实际格式则使用实际格式，否则使用默认格式
            var step54StartTime = System.Diagnostics.Stopwatch.StartNew();
            List<Dictionary<string, object>> subtypeFormatMapping = new List<Dictionary<string, object>>();
            
            // 先处理有实际格式的标题类型
            foreach (var kvp in formatListBySubtype)
            {
                string detailedSubtype = kvp.Key;
                List<Dictionary<string, object>> formats = kvp.Value;
                
                var modeStartTime = System.Diagnostics.Stopwatch.StartNew();
                
                // 过滤掉与正文格式一致的格式
                List<Dictionary<string, object>> filteredFormats = formats.Where(format => !IsFormatEqual(format, bodyModeFormat)).ToList();
                
                Dictionary<string, object> modeFormat;
                if (filteredFormats.Count > 0)
                {
                    // 有与正文不一致的格式，计算这些格式的众数
                    modeFormat = CalculateModeFormat(filteredFormats);
                    System.Diagnostics.Debug.WriteLine($"[步骤5.4]   detailed_subtype {detailedSubtype}: 使用实际格式（过滤后剩余 {filteredFormats.Count} 个格式样本，排除了 {formats.Count - filteredFormats.Count} 个与正文格式一致的样本）");
                }
                else
                {
                    // 所有格式都与正文一致，使用正文格式（文档中的实际格式）
                    modeFormat = new Dictionary<string, object>(bodyModeFormat);
                    System.Diagnostics.Debug.WriteLine($"[步骤5.4]   detailed_subtype {detailedSubtype}: 所有格式都与正文格式一致，使用正文格式（文档实际格式）");
                }
                
                modeStartTime.Stop();
                
                // 从 sentencesByDetailedSubtype 中获取第一个句子的分类信息
                string type = "heading";
                string subtype = "";
                string language = "";
                if (sentencesByDetailedSubtype.ContainsKey(detailedSubtype) && sentencesByDetailedSubtype[detailedSubtype].Count > 0)
                {
                    var firstSentencePair = sentencesByDetailedSubtype[detailedSubtype][0];
                    var firstSentence = firstSentencePair.Value;  // KeyValuePair的Value是Dictionary<string, object>
                    type = firstSentence.ContainsKey("type") ? firstSentence["type"]?.ToString() ?? "heading" : "heading";
                    subtype = firstSentence.ContainsKey("subtype") ? firstSentence["subtype"]?.ToString() ?? "" : "";
                    language = firstSentence.ContainsKey("language") ? firstSentence["language"]?.ToString() ?? "" : "";
                }
                
                Dictionary<string, object> subtypeFormat = new Dictionary<string, object>
                {
                    { "type", type },
                    { "subtype", subtype },
                    { "detailed_subtype", detailedSubtype },
                    { "language", language },
                    { "format_origin", "extracted" },
                    { "format", modeFormat }
                };
                
                subtypeFormatMapping.Add(subtypeFormat);
                System.Diagnostics.Debug.WriteLine($"[步骤5.4]   detailed_subtype {detailedSubtype} 计算众数格式耗时: {modeStartTime.ElapsedMilliseconds}ms（原始 {formats.Count} 个样本）");
            }
            
            // 处理没有实际格式的标题类型（使用默认格式）
            foreach (var kvp in defaultFormatsBySubtype)
            {
                string detailedSubtype = kvp.Key;
                
                // 如果已经有实际格式，跳过
                if (formatListBySubtype.ContainsKey(detailedSubtype))
                {
                    continue;
                }
                
                Dictionary<string, object> defaultFormat = kvp.Value;
                
                // 根据 detailed_subtype 推断 type, subtype, language
                string type = "heading";
                string subtype = "";
                string language = "";
                
                if (detailedSubtype.StartsWith("chinese_"))
                {
                    language = "chinese";
                    if (detailedSubtype.Contains("chapter"))
                    {
                        subtype = "chapter";
                    }
                    else if (detailedSubtype.Contains("keyword"))
                    {
                        subtype = "keyword";
                    }
                    else if (detailedSubtype.Contains("numbered"))
                    {
                        subtype = "numbered";
                    }
                }
                else if (detailedSubtype.StartsWith("english_"))
                {
                    language = "english";
                    if (detailedSubtype.Contains("chapter"))
                    {
                        subtype = "chapter";
                    }
                    else if (detailedSubtype.Contains("keyword"))
                    {
                        subtype = "keyword";
                    }
                    else if (detailedSubtype.Contains("numbered"))
                    {
                        subtype = "numbered";
                    }
                }
                else if (detailedSubtype.StartsWith("mixed_"))
                {
                    language = "mixed";
                    subtype = "numbered";
                }
                else if (detailedSubtype == "markdown_heading")
                {
                    language = "markdown";
                    subtype = "markdown";
                }
                
                Dictionary<string, object> subtypeFormat = new Dictionary<string, object>
                {
                    { "type", type },
                    { "subtype", subtype },
                    { "detailed_subtype", detailedSubtype },
                    { "language", language },
                    { "format_origin", "inferred" },
                    { "format", defaultFormat }
                };
                
                subtypeFormatMapping.Add(subtypeFormat);
                System.Diagnostics.Debug.WriteLine($"[步骤5.4]   detailed_subtype {detailedSubtype}: 使用默认格式（文档中无此类型标题）");
            }
            
            step54StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5.4] ✅ 共处理了 {subtypeFormatMapping.Count} 个detailed_subtype的格式（{formatListBySubtype.Count} 个有实际格式，{subtypeFormatMapping.Count - formatListBySubtype.Count} 个使用默认格式），总耗时: {step54StartTime.ElapsedMilliseconds}ms");
            
            // 添加正文格式到输出列表（有正文用提取的众数，无正文用默认：11pt、空字体名、非加粗）
            Dictionary<string, object> bodyFormatDict = bodyModeFormat.Count > 0
                ? bodyModeFormat
                : GetDefaultBodyFormat();
            string bodyFormatOrigin = bodyModeFormat.Count > 0 ? "extracted" : "inferred";
            Dictionary<string, object> bodyFormat = new Dictionary<string, object>
            {
                { "type", "body" },
                { "subtype", "" },
                { "detailed_subtype", "body" },
                { "language", "" },
                { "format_origin", bodyFormatOrigin },
                { "format", bodyFormatDict }
            };
            subtypeFormatMapping.Add(bodyFormat);
            System.Diagnostics.Debug.WriteLine(bodyModeFormat.Count > 0
                ? "[步骤5.3] ✅ 已添加正文格式到输出列表"
                : "[步骤5.3] ✅ 文档无正文，已添加默认正文格式（11pt、空字体名、非加粗）到输出列表");
            
            step5StartTime.Stop();
            System.Diagnostics.Debug.WriteLine($"[步骤5] ✅ 步骤5完成，总耗时: {step5StartTime.ElapsedMilliseconds}ms ({step5StartTime.ElapsedMilliseconds / 1000.0:F2}秒)");
            System.Diagnostics.Debug.WriteLine($"[步骤5]   时间分布: 提取正文 {step50StartTime.ElapsedMilliseconds}ms, 生成默认格式 {step51StartTime.ElapsedMilliseconds}ms, 分组 {step52StartTime.ElapsedMilliseconds}ms, 搜索和提取 {step53StartTime.ElapsedMilliseconds}ms, 计算众数 {step54StartTime.ElapsedMilliseconds}ms");

            // 步骤6：调用服务端推断各 detailed_subtype 的大纲层级，写入映射项的 level（与 POST /heading_level/infer 约定一致）
            MergeHeadingLevelsFromApi(classification, subtypeFormatMapping);

            return (classification, subtypeFormatMapping);
        }

        /// <summary>
        /// P2/P5/P6/P9：从句级候选池抽样；不回退 readText；不写 level。
        /// </summary>
        private static (DocumentLineClassificationResult classification, List<Dictionary<string, object>> subtypeFormatMapping) ExtractFromSentencePool(
            Word.Document doc,
            List<SampleSentence> sentencePool)
        {
            var emptyHeadings = new Dictionary<string, List<Dictionary<string, object>>>();
            var emptyClass = new DocumentLineClassificationResult(emptyHeadings, new List<Dictionary<string, object>>());
            if (doc == null)
            {
                return (emptyClass, new List<Dictionary<string, object>>());
            }

            if (sentencePool == null || sentencePool.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("⚠️ [sampling_pool] empty_pool: format 使用推断/默认 body 与标题 subtype");
            }

            var recognizer = new HeadingRecognizer();
            var bodyCandidates = new List<Dictionary<string, object>>();
            var sentencesByDetailedSubtype = new Dictionary<string, List<KeyValuePair<string, Dictionary<string, object>>>>();
            var headingsByLanguage = new Dictionary<string, List<Dictionary<string, object>>>();
            int headingOrdinal = 0;
            int poolLine = 0;

            foreach (SampleSentence item in sentencePool ?? new List<SampleSentence>())
            {
                string text = item?.Text ?? "";
                string trimmed = text.Trim();
                if (string.IsNullOrEmpty(trimmed))
                {
                    continue;
                }

                poolLine++;
                Tuple<bool, string, string> headingResult = recognizer.IsHeading(text.TrimEnd());
                if (headingResult.Item1)
                {
                    string langType = headingResult.Item2 ?? "";
                    string headingType = headingResult.Item3 ?? "";
                    string detailedSubtype = recognizer.GetDetailedSubtype(trimmed, langType, headingType);
                    var headingInfo = new Dictionary<string, object>
                    {
                        { "text", text },
                        { "language", langType },
                        { "subtype", headingType },
                        { "detailed_subtype", detailedSubtype },
                        { "type", "heading" },
                        { "name", item.Name ?? "" },
                        { "line", poolLine }
                    };

                    if (!sentencesByDetailedSubtype.ContainsKey(detailedSubtype))
                    {
                        sentencesByDetailedSubtype[detailedSubtype] = new List<KeyValuePair<string, Dictionary<string, object>>>();
                    }

                    string syntheticKey = "H_" + headingOrdinal++;
                    sentencesByDetailedSubtype[detailedSubtype].Add(new KeyValuePair<string, Dictionary<string, object>>(syntheticKey, headingInfo));

                    if (!headingsByLanguage.ContainsKey(langType))
                    {
                        headingsByLanguage[langType] = new List<Dictionary<string, object>>();
                    }

                    headingsByLanguage[langType].Add(headingInfo);
                }
                else
                {
                    bodyCandidates.Add(new Dictionary<string, object>
                    {
                        { "text", trimmed },
                        { "name", item.Name ?? "" }
                    });
                }
            }

            var random = new Random();
            var bodyFormats = new List<Dictionary<string, object>>();
            List<Dictionary<string, object>> sampledBody = bodyCandidates;
            if (sampledBody.Count > MaxBodySamplesForFormat)
            {
                sampledBody = bodyCandidates.OrderBy(_ => random.Next()).Take(MaxBodySamplesForFormat).ToList();
            }

            foreach (Dictionary<string, object> bodyInfo in sampledBody)
            {
                string sentenceText = bodyInfo.ContainsKey("text") ? bodyInfo["text"]?.ToString() ?? "" : "";
                if (string.IsNullOrEmpty(sentenceText))
                {
                    continue;
                }

                foreach (Word.Range range in FindAllSentenceMatches(doc, sentenceText))
                {
                    bodyFormats.Add(ExtractFormatFromRange(range));
                }
            }

            if (bodyCandidates.Count == 0 && sentencePool != null && sentencePool.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine("⚠️ [sampling_pool] no_body_candidates_in_pool");
            }

            Dictionary<string, object> bodyModeFormat = bodyFormats.Count > 0 ? CalculateModeFormat(bodyFormats) : new Dictionary<string, object>();

            Dictionary<string, Dictionary<string, object>> defaultFormatsBySubtype = new Dictionary<string, Dictionary<string, object>>();
            foreach (string detailedSubtype in GetAllPossibleDetailedSubtypes())
            {
                defaultFormatsBySubtype[detailedSubtype] = GenerateDefaultFormat(bodyModeFormat, detailedSubtype);
            }

            var formatListBySubtype = new Dictionary<string, List<Dictionary<string, object>>>();
            foreach (var kvp in sentencesByDetailedSubtype)
            {
                string detailedSubtype = kvp.Key;
                List<KeyValuePair<string, Dictionary<string, object>>> toSample = kvp.Value;
                if (toSample.Count > MaxHeadingSamplesPerSubtype)
                {
                    toSample = toSample.OrderBy(_ => random.Next()).Take(MaxHeadingSamplesPerSubtype).ToList();
                }

                var formatsFor = new List<Dictionary<string, object>>();
                foreach (KeyValuePair<string, Dictionary<string, object>> pair in toSample)
                {
                    string sentenceText = pair.Value.ContainsKey("text") ? pair.Value["text"]?.ToString() ?? "" : "";
                    if (string.IsNullOrEmpty(sentenceText))
                    {
                        continue;
                    }

                    foreach (Word.Range range in FindAllSentenceMatches(doc, sentenceText))
                    {
                        formatsFor.Add(ExtractFormatFromRange(range));
                    }
                }

                if (formatsFor.Count > 0)
                {
                    formatListBySubtype[detailedSubtype] = formatsFor;
                }
            }

            List<Dictionary<string, object>> subtypeFormatMapping = BuildSubtypeFormatMappingList(
                bodyModeFormat,
                bodyFormats,
                formatListBySubtype,
                sentencesByDetailedSubtype,
                defaultFormatsBySubtype);

            var classification = new DocumentLineClassificationResult(headingsByLanguage, new List<Dictionary<string, object>>());
            return (classification, subtypeFormatMapping);
        }

        private static List<Dictionary<string, object>> BuildSubtypeFormatMappingList(
            Dictionary<string, object> bodyModeFormat,
            List<Dictionary<string, object>> bodyFormats,
            Dictionary<string, List<Dictionary<string, object>>> formatListBySubtype,
            Dictionary<string, List<KeyValuePair<string, Dictionary<string, object>>>> sentencesByDetailedSubtype,
            Dictionary<string, Dictionary<string, object>> defaultFormatsBySubtype)
        {
            var subtypeFormatMapping = new List<Dictionary<string, object>>();

            foreach (var kvp in formatListBySubtype)
            {
                string detailedSubtype = kvp.Key;
                List<Dictionary<string, object>> formats = kvp.Value;
                List<Dictionary<string, object>> filteredFormats = formats.Where(format => !IsFormatEqual(format, bodyModeFormat)).ToList();
                Dictionary<string, object> modeFormat;
                string formatOrigin;
                if (filteredFormats.Count > 0)
                {
                    modeFormat = CalculateModeFormat(filteredFormats);
                    formatOrigin = "extracted";
                }
                else
                {
                    modeFormat = bodyModeFormat.Count > 0
                        ? new Dictionary<string, object>(bodyModeFormat)
                        : GetDefaultBodyFormat();
                    formatOrigin = bodyModeFormat.Count > 0 ? "extracted" : "inferred";
                }

                string type = "heading";
                string subtype = "";
                string language = "";
                if (sentencesByDetailedSubtype.ContainsKey(detailedSubtype) && sentencesByDetailedSubtype[detailedSubtype].Count > 0)
                {
                    Dictionary<string, object> firstSentence = sentencesByDetailedSubtype[detailedSubtype][0].Value;
                    type = firstSentence.ContainsKey("type") ? firstSentence["type"]?.ToString() ?? "heading" : "heading";
                    subtype = firstSentence.ContainsKey("subtype") ? firstSentence["subtype"]?.ToString() ?? "" : "";
                    language = firstSentence.ContainsKey("language") ? firstSentence["language"]?.ToString() ?? "" : "";
                }

                subtypeFormatMapping.Add(new Dictionary<string, object>
                {
                    { "type", type },
                    { "subtype", subtype },
                    { "detailed_subtype", detailedSubtype },
                    { "language", language },
                    { "format_origin", formatOrigin },
                    { "format", modeFormat }
                });
            }

            foreach (var kvp in defaultFormatsBySubtype)
            {
                string detailedSubtype = kvp.Key;
                if (formatListBySubtype.ContainsKey(detailedSubtype))
                {
                    continue;
                }

                InferSubtypeFields(detailedSubtype, out string type, out string subtype, out string language);
                subtypeFormatMapping.Add(new Dictionary<string, object>
                {
                    { "type", type },
                    { "subtype", subtype },
                    { "detailed_subtype", detailedSubtype },
                    { "language", language },
                    { "format_origin", "inferred" },
                    { "format", kvp.Value }
                });
            }

            Dictionary<string, object> bodyFormatDict = bodyModeFormat.Count > 0 ? bodyModeFormat : GetDefaultBodyFormat();
            string bodyFormatOrigin = bodyModeFormat.Count > 0 ? "extracted" : "inferred";
            subtypeFormatMapping.Add(new Dictionary<string, object>
            {
                { "type", "body" },
                { "subtype", "" },
                { "detailed_subtype", "body" },
                { "language", "" },
                { "format_origin", bodyFormatOrigin },
                { "format", bodyFormatDict }
            });

            return subtypeFormatMapping;
        }

        private static void InferSubtypeFields(string detailedSubtype, out string type, out string subtype, out string language)
        {
            type = "heading";
            subtype = "";
            language = "";
            if (detailedSubtype.StartsWith("chinese_"))
            {
                language = "chinese";
                if (detailedSubtype.Contains("chapter"))
                {
                    subtype = "chapter";
                }
                else if (detailedSubtype.Contains("keyword"))
                {
                    subtype = "keyword";
                }
                else if (detailedSubtype.Contains("numbered"))
                {
                    subtype = "numbered";
                }
            }
            else if (detailedSubtype.StartsWith("english_"))
            {
                language = "english";
                if (detailedSubtype.Contains("chapter"))
                {
                    subtype = "chapter";
                }
                else if (detailedSubtype.Contains("keyword"))
                {
                    subtype = "keyword";
                }
                else if (detailedSubtype.Contains("numbered"))
                {
                    subtype = "numbered";
                }
            }
            else if (detailedSubtype.StartsWith("mixed_"))
            {
                language = "mixed";
                subtype = "numbered";
            }
            else if (detailedSubtype == "markdown_heading")
            {
                language = "markdown";
                subtype = "markdown";
            }
        }

        /// <summary>
        /// 与 Python <c>format_headings_block_for_prompt</c> 对齐：按原文顺序的标题块文本。
        /// </summary>
        private static string FormatHeadingsBlockForPrompt(List<Dictionary<string, object>> headingsOrdered)
        {
            if (headingsOrdered == null || headingsOrdered.Count == 0)
            {
                return "";
            }

            var sb = new StringBuilder();
            for (int i = 0; i < headingsOrdered.Count; i++)
            {
                Dictionary<string, object> h = headingsOrdered[i];
                int lineNum = GetHeadingLineNumberForPrompt(h);
                string dst = h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                string text = h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                text = text.Trim();
                sb.Append($"{i + 1}. 行{lineNum}  [{dst}]  {text}");
                if (i < headingsOrdered.Count - 1)
                {
                    sb.Append("\n");
                }
            }

            return sb.ToString();
        }

        private static int GetHeadingLineNumberForPrompt(Dictionary<string, object> h)
        {
            if (h != null && h.TryGetValue("line", out object o) && o != null)
            {
                if (o is int ii)
                {
                    return ii;
                }

                if (int.TryParse(o.ToString(), out int p))
                {
                    return p;
                }
            }

            return 0;
        }

        /// <summary>
        /// 各 <c>detailed_subtype</c> 在文中首次出现的顺序（不重复）。
        /// </summary>
        private static List<string> OrderedUniqueDetailedSubtypes(List<Dictionary<string, object>> headingsOrdered)
        {
            var list = new List<string>();
            if (headingsOrdered == null)
            {
                return list;
            }

            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (Dictionary<string, object> h in headingsOrdered)
            {
                string dst = h != null && h.ContainsKey("detailed_subtype")
                    ? h["detailed_subtype"]?.ToString() ?? ""
                    : "";
                if (string.IsNullOrEmpty(dst) || !seen.Add(dst))
                {
                    continue;
                }

                list.Add(dst);
            }

            return list;
        }

        /// <summary>
        /// 将接口返回的 detailed_subtype → level 合并到 subtype 格式列表。
        /// 仅对文档中确有标题样本的项（<c>format_origin == "extracted"</c>）写入 level；默认补全的推断项不写入。
        /// </summary>
        private static void ApplyHeadingLevelsToSubtypeFormatMapping(
            List<Dictionary<string, object>> subtypeFormatMapping,
            Dictionary<string, int> levelsBySubtype)
        {
            if (subtypeFormatMapping == null || levelsBySubtype == null || levelsBySubtype.Count == 0)
            {
                return;
            }

            foreach (Dictionary<string, object> item in subtypeFormatMapping)
            {
                string t = item.ContainsKey("type") ? item["type"]?.ToString() ?? "" : "";
                if (t != "heading")
                {
                    continue;
                }

                string origin = item.ContainsKey("format_origin") ? item["format_origin"]?.ToString() ?? "" : "";
                if (origin != "extracted")
                {
                    continue;
                }

                string dst = item.ContainsKey("detailed_subtype") ? item["detailed_subtype"]?.ToString() ?? "" : "";
                if (string.IsNullOrEmpty(dst))
                {
                    continue;
                }

                if (levelsBySubtype.TryGetValue(dst, out int level))
                {
                    item["level"] = level;
                }
            }
        }

        /// <summary>
        /// 请求 /heading_level/infer（请求体仅含文中真实标题）并将 level 写入映射中对应「提取」项（失败则保持原样）。
        /// </summary>
        private static void MergeHeadingLevelsFromApi(
            DocumentLineClassificationResult classification,
            List<Dictionary<string, object>> subtypeFormatMapping)
        {
            try
            {
                List<Dictionary<string, object>> headingsOrdered = FlattenHeadingsInDocumentOrder(classification);
                string headingsBlock = FormatHeadingsBlockForPrompt(headingsOrdered);
                List<string> uniqueTypes = OrderedUniqueDetailedSubtypes(headingsOrdered);
                if (string.IsNullOrWhiteSpace(headingsBlock) || uniqueTypes.Count == 0)
                {
                    System.Diagnostics.Debug.WriteLine("[FormatExtractor] 标题块为空或无类型，跳过层级推断");
                    return;
                }

                Dictionary<string, int> levels = HeadingLevelApiClient.TryInferHeadingLevels(headingsBlock, uniqueTypes);
                ApplyHeadingLevelsToSubtypeFormatMapping(subtypeFormatMapping, levels);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FormatExtractor] 合并标题层级失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取所有可能的detailed_subtype列表
        /// </summary>
        private static List<string> GetAllPossibleDetailedSubtypes()
        {
            return new List<string>
            {
                // 中文类型
                "chinese_chapter",
                "chinese_keyword",
                "chinese_numbered_level3",
                "chinese_numbered_level2",
                "chinese_numbered_bracketed_chinese",
                "chinese_numbered_bracketed_arabic",
                "chinese_numbered_chinese_digit",
                "chinese_numbered_arabic_digit",
                "chinese_numbered_circle",
                
                // 英文类型
                "english_chapter",
                "english_keyword",
                "english_numbered_level3",
                "english_numbered_level2",
                "english_numbered_roman",
                "english_numbered_uppercase",
                "english_numbered_lowercase",
                "english_numbered_bracketed",
                "english_numbered_arabic",
                
                // 混合类型
                "mixed_numbered",
                
                // Markdown类型
                "markdown_heading"
            };
        }
        
        /// <summary>
        /// 根据detailed_subtype获取标题层级（1=最高级，2=中级，3=低级）
        /// </summary>
        private static int GetHeadingLevel(string detailedSubtype)
        {
            // Level 1（最高级）：chapter, keyword, markdown_heading, 一级序号
            if (detailedSubtype.Contains("chapter") || 
                detailedSubtype.Contains("keyword") || 
                detailedSubtype == "markdown_heading" ||
                detailedSubtype == "chinese_numbered_arabic_digit" ||
                detailedSubtype == "chinese_numbered_chinese_digit" ||
                detailedSubtype == "chinese_numbered_bracketed_chinese" ||
                detailedSubtype == "chinese_numbered_bracketed_arabic" ||
                detailedSubtype == "chinese_numbered_circle" ||
                detailedSubtype == "english_numbered_arabic" ||
                detailedSubtype == "english_numbered_roman" ||
                detailedSubtype == "english_numbered_uppercase" ||
                detailedSubtype == "english_numbered_lowercase" ||
                detailedSubtype == "english_numbered_bracketed" ||
                detailedSubtype == "mixed_numbered")
            {
                return 1;
            }
            
            // Level 2（中级）：二级序号
            if (detailedSubtype.Contains("level2"))
            {
                return 2;
            }
            
            // Level 3（低级）：三级序号
            if (detailedSubtype.Contains("level3"))
            {
                return 3;
            }
            
            // 默认返回 Level 1
            return 1;
        }
        
        /// <summary>
        /// 无正文时使用的默认正文格式：字号 11、空字体名、非加粗
        /// </summary>
        private static Dictionary<string, object> GetDefaultBodyFormat()
        {
            return new Dictionary<string, object>
            {
                { "font_name", "" },
                { "font_size", 11f },
                { "font_color", "automatic" },
                { "background_color", "automatic" },
                { "is_bold", false },
                { "has_underline", false },
                { "has_strikethrough", false }
            };
        }
        
        /// <summary>
        /// 根据正文格式和标题类型生成默认格式
        /// 
        /// 规则说明：
        /// 1. 字体名称：与正文保持一致
        /// 2. 字体大小：与正文一致（Level 1/2/3 均不改变字号）
        /// 3. 字体颜色：自动（automatic）
        /// 4. 背景颜色：自动（automatic）
        /// 5. 是否加粗：根据层级决定
        ///    - Level 1 和 Level 2：加粗（true）
        ///    - Level 3：不加粗（false）
        /// 6. 下划线：无（false）
        /// 7. 删除线：无（false）
        /// </summary>
        /// <param name="bodyFormat">正文格式</param>
        /// <param name="detailedSubtype">标题的detailed_subtype</param>
        /// <returns>默认格式字典</returns>
        private static Dictionary<string, object> GenerateDefaultFormat(Dictionary<string, object> bodyFormat, string detailedSubtype)
        {
            Dictionary<string, object> defaultFormat = new Dictionary<string, object>();
            
            // 获取正文格式的各个属性
            string fontName = bodyFormat.ContainsKey("font_name") ? bodyFormat["font_name"]?.ToString() ?? "" : "";
            float bodyFontSize = bodyFormat.ContainsKey("font_size") ? 
                (bodyFormat["font_size"] is float f ? f : (float.TryParse(bodyFormat["font_size"]?.ToString(), out float fs) ? fs : 11f)) : 11f;
            
            // 获取标题层级
            int level = GetHeadingLevel(detailedSubtype);
            
            // 字体名称：与正文保持一致
            defaultFormat["font_name"] = fontName;
            
            // 字体大小：与正文一致，各层级均不改变
            defaultFormat["font_size"] = bodyFontSize;
            
            // 字体颜色：自动
            defaultFormat["font_color"] = "automatic";
            
            // 背景颜色：自动
            defaultFormat["background_color"] = "automatic";
            
            // 是否加粗：根据层级决定
            // Level 1 和 Level 2 加粗，Level 3 不加粗
            if (level == 1 || level == 2)
            {
                defaultFormat["is_bold"] = true; // Level 1 和 Level 2 加粗
            }
            else
            {
                defaultFormat["is_bold"] = false; // Level 3 不加粗
            }
            
            // 是否有下划线：无
            defaultFormat["has_underline"] = false;
            
            // 是否有删除线：无
            defaultFormat["has_strikethrough"] = false;
            
            return defaultFormat;
        }
        
        /// <summary>
        /// 在Word文档中查找所有匹配指定文本的Range位置
        /// </summary>
        private static List<Word.Range> FindAllSentenceMatches(Word.Document doc, string sentenceText)
        {
            List<Word.Range> matches = new List<Word.Range>();
            
            try
            {
                if (doc == null || string.IsNullOrEmpty(sentenceText))
                {
                    return matches;
                }
                
                // 转换换行符为Word代码格式
                string convertedText = WordRangeFinder.ConvertNewlinesToWordCodes(sentenceText);
                
                // 使用WordRangeFinder的FindAllMatches方法
                // 注意：FindAllMatches是private的，我们需要自己实现
                int searchStart = 0;
                int lastFoundStart = -1;
                int lastFoundEnd = -1;
                int maxIterations = 1000;
                int iterationCount = 0;
                
                while (searchStart < doc.Content.End && iterationCount < maxIterations)
                {
                    iterationCount++;
                    
                    Word.Range searchRange = doc.Range(searchStart, doc.Content.End);
                    searchRange.Find.ClearFormatting();
                    searchRange.Find.Text = convertedText;
                    searchRange.Find.MatchCase = true;
                    searchRange.Find.MatchWholeWord = false;
                    searchRange.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    
                    if (searchRange.Find.Execute())
                    {
                        // 检查是否找到了新的匹配
                        if (searchRange.Start == lastFoundStart && searchRange.End == lastFoundEnd)
                        {
                            break;
                        }
                        
                        if (searchRange.Start < searchStart)
                        {
                            break;
                        }
                        
                        Word.Range foundRange = doc.Range(searchRange.Start, searchRange.End);
                        matches.Add(foundRange);
                        
                        lastFoundStart = searchRange.Start;
                        lastFoundEnd = searchRange.End;
                        
                        int newSearchStart = searchRange.End + 1;
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
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[FindAllSentenceMatches] ❌ 异常: {ex.Message}");
            }
            
            return matches;
        }
        
        /// <summary>
        /// 从Word Range提取格式信息
        /// </summary>
        internal static Dictionary<string, object> ExtractFormatFromRange(Word.Range range)
        {
            Dictionary<string, object> format = new Dictionary<string, object>();
            
            try
            {
                if (range == null)
                {
                    return format;
                }
                
                // 字体名称
                string fontName = range.Font?.Name ?? "";
                format["font_name"] = fontName;
                
                // 字体大小
                float fontSize = range.Font?.Size ?? 0f;
                format["font_size"] = fontSize;
                
                // 字体颜色（转换为RGB十六进制字符串）
                string fontColor = ConvertWdColorToRgb(range.Font?.Color ?? Word.WdColor.wdColorAutomatic);
                format["font_color"] = fontColor;
                
                // 背景颜色（转换为RGB十六进制字符串）
                string backgroundColor = ConvertWdColorToRgb(range.Shading?.BackgroundPatternColor ?? Word.WdColor.wdColorAutomatic);
                format["background_color"] = backgroundColor;
                
                // 是否加粗
                bool isBold = (range.Font?.Bold ?? 0) != 0;
                format["is_bold"] = isBold;
                
                // 是否有下划线
                bool hasUnderline = (range.Font?.Underline ?? Word.WdUnderline.wdUnderlineNone) != Word.WdUnderline.wdUnderlineNone;
                format["has_underline"] = hasUnderline;
                
                // 是否有删除线
                bool hasStrikethrough = (range.Font?.StrikeThrough ?? 0) != 0;
                format["has_strikethrough"] = hasStrikethrough;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractFormatFromRange] ❌ 异常: {ex.Message}");
            }
            
            return format;
        }
        
        /// <summary>
        /// 将Word.WdColor转换为RGB十六进制字符串（如 #RRGGBB）
        /// </summary>
        private static string ConvertWdColorToRgb(Word.WdColor wdColor)
        {
            try
            {
                // Word.WdColor.wdColorAutomatic 的值是 -16777216
                if (wdColor == Word.WdColor.wdColorAutomatic || (int)wdColor == -16777216)
                {
                    return "automatic";
                }
                
                // Word颜色值是BGR格式，需要转换为RGB
                int colorValue = (int)wdColor;
                
                // 提取RGB分量（BGR格式）
                int b = (colorValue & 0xFF0000) >> 16;
                int g = (colorValue & 0x00FF00) >> 8;
                int r = colorValue & 0x0000FF;
                
                // 转换为RGB十六进制字符串
                return $"#{r:X2}{g:X2}{b:X2}";
            }
            catch
            {
                return "unknown";
            }
        }
        
        /// <summary>
        /// 判断两个格式是否相等
        /// </summary>
        private static bool IsFormatEqual(Dictionary<string, object> format1, Dictionary<string, object> format2)
        {
            if (format1 == null || format2 == null)
            {
                return format1 == format2;
            }
            
            if (format1.Count != format2.Count)
            {
                return false;
            }
            
            string[] formatKeys = { "font_name", "font_size", "font_color", "background_color", "is_bold", "has_underline", "has_strikethrough" };
            
            foreach (string key in formatKeys)
            {
                object value1 = format1.ContainsKey(key) ? format1[key] : null;
                object value2 = format2.ContainsKey(key) ? format2[key] : null;
                
                // 处理浮点数比较（字体大小）
                if (key == "font_size")
                {
                    float size1 = value1 is float f1 ? f1 : (value1 != null && float.TryParse(value1.ToString(), out float s1) ? s1 : 0f);
                    float size2 = value2 is float f2 ? f2 : (value2 != null && float.TryParse(value2.ToString(), out float s2) ? s2 : 0f);
                    if (Math.Abs(size1 - size2) > 0.01f) // 允许0.01的误差
                    {
                        return false;
                    }
                }
                // 处理布尔值比较
                else if (key == "is_bold" || key == "has_underline" || key == "has_strikethrough")
                {
                    bool bool1 = value1 is bool b1 ? b1 : (value1?.ToString() == "True" || value1?.ToString() == "true");
                    bool bool2 = value2 is bool b2 ? b2 : (value2?.ToString() == "True" || value2?.ToString() == "true");
                    if (bool1 != bool2)
                    {
                        return false;
                    }
                }
                // 处理字符串比较
                else
                {
                    string str1 = value1?.ToString() ?? "";
                    string str2 = value2?.ToString() ?? "";
                    if (str1 != str2)
                    {
                        return false;
                    }
                }
            }
            
            return true;
        }
        
        /// <summary>
        /// 计算格式列表的众数格式
        /// </summary>
        internal static Dictionary<string, object> CalculateModeFormat(List<Dictionary<string, object>> formats)
        {
            Dictionary<string, object> modeFormat = new Dictionary<string, object>();
            
            if (formats == null || formats.Count == 0)
            {
                return modeFormat;
            }
            
            // 对每个格式属性计算众数
            string[] formatKeys = { "font_name", "font_size", "font_color", "background_color", "is_bold", "has_underline", "has_strikethrough" };
            
            foreach (string key in formatKeys)
            {
                // 统计每个值的出现次数
                Dictionary<string, int> valueCounts = new Dictionary<string, int>();
                
                foreach (var format in formats)
                {
                    if (format.ContainsKey(key))
                    {
                        string value = format[key]?.ToString() ?? "";
                        if (!valueCounts.ContainsKey(value))
                        {
                            valueCounts[value] = 0;
                        }
                        valueCounts[value]++;
                    }
                }
                
                // 找出出现次数最多的值（众数）
                if (valueCounts.Count > 0)
                {
                    string modeValue = valueCounts.OrderByDescending(kvp => kvp.Value).First().Key;
                    
                    // 根据key的类型转换值
                    if (key == "font_size")
                    {
                        if (float.TryParse(modeValue, out float fontSize))
                        {
                            modeFormat[key] = fontSize;
                        }
                        else
                        {
                            modeFormat[key] = 0f;
                        }
                    }
                    else if (key == "is_bold" || key == "has_underline" || key == "has_strikethrough")
                    {
                        // 布尔值：统计True和False的数量，取众数
                        int trueCount = formats.Count(f => f.ContainsKey(key) && (f[key] is bool b && b || f[key]?.ToString() == "True" || f[key]?.ToString() == "true"));
                        int falseCount = formats.Count - trueCount;
                        modeFormat[key] = trueCount >= falseCount;
                    }
                    else
                    {
                        modeFormat[key] = modeValue;
                    }
                }
                else
                {
                    // 如果没有找到值，设置默认值
                    if (key == "font_size")
                    {
                        modeFormat[key] = 0f;
                    }
                    else if (key == "is_bold" || key == "has_underline" || key == "has_strikethrough")
                    {
                        modeFormat[key] = false;
                    }
                    else
                    {
                        modeFormat[key] = "";
                    }
                }
            }
            
            return modeFormat;
        }
    }
}

