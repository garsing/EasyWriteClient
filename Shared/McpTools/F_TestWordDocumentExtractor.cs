using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 测试Word文档提取器工具
    /// 用于测试WordDocumentExtractor的功能，提取当前文档内容并打印到控制台
    /// </summary>
    public static class F_TestWordDocumentExtractor
    {
        /// <summary>
        /// 注册测试Word文档提取器工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_test_word_document_extractor"] = async (args) =>
            {
                try
                {
                    if (wordApplication == null)
                    {
                        return new ToolResult { Success = false, Error = "Word应用程序实例不可用" };
                    }

                    dynamic wordApp = wordApplication;

                    if (wordApp.ActiveDocument == null)
                    {
                        return new ToolResult { Success = false, Error = "没有活动的Word文档" };
                    }

                    Word.Document doc = wordApp.ActiveDocument;

                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("=== F_test_word_document_extractor 工具执行 ===");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine($"文档名称: {doc.Name ?? "未命名文档"}");
                    System.Diagnostics.Debug.WriteLine("");

                    // 同时输出到控制台（Console）
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("=== F_test_word_document_extractor 工具执行 ===");
                    Console.WriteLine("==========================================");
                    Console.WriteLine($"文档名称: {doc.Name ?? "未命名文档"}");
                    Console.WriteLine("");

                    // 调用WordDocumentExtractor执行完整流程（从文件读入到生成快照）
                    var result = WordDocumentExtractor.ProcessDocument(
                        doc,
                        new ProcessDocumentOptions { VerboseDebug = true });

                    // 打印处理结果
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("=== 处理结果汇总 ===");
                    Console.WriteLine("==========================================");
                    
                    // 打印处理后的文本（已转义）
                    string escapedProcessedText = TextUtils.EscapeNewlinesForDb(result.ReadText);
                    Console.WriteLine($"处理后的文本长度: {result.ReadText.Length} 字符");
                    Console.WriteLine($"处理后的文本（已转义，前200字符）: {escapedProcessedText.Substring(0, Math.Min(200, escapedProcessedText.Length))}...");
                    
                    // 打印分块信息
                    Console.WriteLine($"分块数量: {result.ChunkInfo.Count}");
                    for (int i = 0; i < result.ChunkInfo.Count; i++)
                    {
                        var chunkInfo = result.ChunkInfo[i];
                        string escapedChunk = TextUtils.EscapeNewlinesForDb(chunkInfo.Content);
                        string escapedDisplayContent = TextUtils.EscapeNewlinesForDb(chunkInfo.DisplayContent);
                        Console.WriteLine($"  Chunk {chunkInfo.ChunkIdx} 长度: {chunkInfo.Content.Length} 字符");
                        Console.WriteLine($"  Chunk {chunkInfo.ChunkIdx} 内容（已转义，前200字符）: {escapedChunk.Substring(0, Math.Min(200, escapedChunk.Length))}...");
                        Console.WriteLine($"  Chunk {chunkInfo.ChunkIdx} Display Content 长度: {chunkInfo.DisplayContent.Length} 字符");
                        Console.WriteLine($"  Chunk {chunkInfo.ChunkIdx} Display Content（已转义，前200字符）: {escapedDisplayContent.Substring(0, Math.Min(200, escapedDisplayContent.Length))}...");
                    }
                    
                    // 打印句子映射信息
                    Console.WriteLine($"句子映射数量: {result.NameToContentMap.Count}");
                    int mappingCount = 0;
                    foreach (var kvp in result.NameToContentMap)
                    {
                        if (mappingCount < 5) // 只打印前5个映射
                        {
                            string escapedContent = TextUtils.EscapeNewlinesForDb(kvp.Value);
                            Console.WriteLine($"  映射 {mappingCount + 1}: {kvp.Key} -> {escapedContent.Substring(0, Math.Min(100, escapedContent.Length))}...");
                        }
                        mappingCount++;
                    }
                    if (result.NameToContentMap.Count > 5)
                    {
                        Console.WriteLine($"  ... 还有 {result.NameToContentMap.Count - 5} 个映射");
                    }
                    
                    // 打印快照信息
                    string escapedSnapshot = TextUtils.EscapeNewlinesForDb(result.Snapshot);
                    string escapedSegSnapshot = TextUtils.EscapeNewlinesForDb(result.SegSnapshot);
                    Console.WriteLine($"快照（snapshot）长度: {result.Snapshot.Length} 字符");
                    Console.WriteLine($"快照（snapshot，已转义）: {escapedSnapshot}");
                    Console.WriteLine($"快照（seg_snapshot）长度: {result.SegSnapshot.Length} 字符");
                    Console.WriteLine($"快照（seg_snapshot，已转义）: {escapedSegSnapshot}");
                    
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");

                    // 测试SegSnapshot中的句子搜索
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("=== SegSnapshot 句子搜索测试 ===");
                    Console.WriteLine("==========================================");
                    var searchTestResult = TestSegSnapshotSearch(doc, result.SegSnapshot, result.NameToContentMap);
                    Console.WriteLine($"测试完成：成功 {searchTestResult.SuccessCount} 个，失败 {searchTestResult.FailCount} 个");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");

                    await Task.CompletedTask;

                    return new ToolResult
                    {
                        Success = true,
                        Data = new Dictionary<string, object>
                        {
                            { "document_name", doc.Name ?? "未命名文档" },
                            { "read_text_length", result.ReadText.Length },
                            { "chunks_count", result.ChunkInfo.Count },
                            { "sentence_mappings_count", result.NameToContentMap.Count },
                            { "snapshot_length", result.Snapshot.Length },
                            { "seg_snapshot_length", result.SegSnapshot.Length },
                            { "search_test_success_count", searchTestResult.SuccessCount },
                            { "search_test_fail_count", searchTestResult.FailCount }
                        }
                    };
                }
                catch (Exception ex)
                {
                    string errorMsg = $"测试Word文档提取器失败: {ex.Message}";
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] {errorMsg}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 堆栈跟踪: {ex.StackTrace}");
                    Console.WriteLine($"[ERROR] {errorMsg}");
                    Console.WriteLine($"[ERROR] 堆栈跟踪: {ex.StackTrace}");
                    return new ToolResult { Success = false, Error = errorMsg };
                }
            };
        }

        /// <summary>
        /// 测试SegSnapshot中的句子搜索
        /// </summary>
        /// <param name="doc">Word文档</param>
        /// <param name="segSnapshot">SegSnapshot字符串，格式：S_00000,S_00001,S_00002;S_00003,S_00004,S_00005,S_00006;S_00007,S_00008,S_00009</param>
        /// <param name="nameToContentMap">句子名称到内容的映射字典</param>
        /// <returns>测试结果</returns>
        private static (int SuccessCount, int FailCount) TestSegSnapshotSearch(
            Word.Document doc, 
            string segSnapshot, 
            Dictionary<string, string> nameToContentMap)
        {
            int successCount = 0;
            int failCount = 0;
            int searchSequenceNumber = 0; // 搜索序号计数器

            if (string.IsNullOrEmpty(segSnapshot))
            {
                Console.WriteLine("SegSnapshot为空，无法进行测试");
                return (0, 0);
            }

            // 解析SegSnapshot：seg与seg之间用;隔开，seg内句子与句子用,逗号隔开
            string[] segs = segSnapshot.Split(';');
            Console.WriteLine($"共解析到 {segs.Length} 个seg");

            for (int segIndex = 0; segIndex < segs.Length; segIndex++)
            {
                string seg = segs[segIndex].Trim();
                if (string.IsNullOrEmpty(seg))
                {
                    continue;
                }

                // 解析seg内的句子名称
                string[] sentenceNames = seg.Split(',');
                List<string> sentenceNameList = new List<string>();
                foreach (var name in sentenceNames)
                {
                    string trimmedName = name.Trim();
                    if (!string.IsNullOrEmpty(trimmedName))
                    {
                        sentenceNameList.Add(trimmedName);
                    }
                }

                if (sentenceNameList.Count == 0)
                {
                    continue;
                }

                Console.WriteLine("");
                Console.WriteLine($"--- Seg {segIndex + 1} (共 {sentenceNameList.Count} 个句子) ---");
                Console.WriteLine($"句子名称: {string.Join(",", sentenceNameList)}");

                // 第一步：单独搜索每个句子
                Console.WriteLine($"第一步：单独搜索每个句子");
                int singleSearchIndex = 0;
                foreach (var sentenceName in sentenceNameList)
                {
                    singleSearchIndex++;
                    searchSequenceNumber++;
                    System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] Seg {segIndex + 1}, 单独搜索 {singleSearchIndex}/{sentenceNameList.Count}: {sentenceName}");
                    
                    if (nameToContentMap.TryGetValue(sentenceName, out string sentenceContent))
                    {
                        System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] 句子内容长度: {sentenceContent.Length}");

                        if (sentenceContent.Length <= 1)
                        {
                            System.Threading.Thread.Sleep(50);
                        }

                        Word.Range foundRange = WordRangeFinder.FindSentenceRangeInDocument(doc, sentenceContent, sentenceName);
                        if (foundRange != null)
                        {
                            successCount++;
                            Console.WriteLine($"  [{searchSequenceNumber}] ✅ {sentenceName}: 搜索成功 (长度: {sentenceContent.Length})");
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ✅ 搜索成功，Range位置: Start={foundRange.Start}, End={foundRange.End}");
                        }
                        else
                        {
                            failCount++;
                            string escapedContent = sentenceContent.Replace("\r", "\\r").Replace("\n", "\\n");
                            Console.WriteLine($"  [{searchSequenceNumber}] ❌ {sentenceName}: 搜索失败 (长度: {sentenceContent.Length})");
                            Console.WriteLine($"     完整搜索内容: {escapedContent}");
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ❌ 搜索失败，完整内容: {escapedContent}");
                        }
                    }
                    else
                    {
                        failCount++;
                        Console.WriteLine($"  [{searchSequenceNumber}] ❌ {sentenceName}: 在映射表中找不到内容");
                        System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ❌ 在映射表中找不到句子名称: {sentenceName}");
                    }
                }

                // 第二步：相邻句子组合搜索（支持多个相邻句子组合，最多10个，引入随机性）
                Console.WriteLine($"第二步：相邻句子组合搜索（随机测试多个相邻句子组合）");
                
                // 生成所有可能的相邻句子组合（长度从2到min(10, seg中句子数量)）
                var allCombinations = new List<(int startIndex, int length)>();
                int maxLength = Math.Min(10, sentenceNameList.Count);
                
                for (int length = 2; length <= maxLength; length++)
                {
                    for (int startIndex = 0; startIndex <= sentenceNameList.Count - length; startIndex++)
                    {
                        allCombinations.Add((startIndex, length));
                    }
                }
                
                // 引入随机性：随机选择一部分组合进行测试
                // 如果组合数量较少（<=20），全部测试；否则随机选择20个组合
                var random = new Random();
                var selectedCombinations = allCombinations;
                if (allCombinations.Count > 20)
                {
                    selectedCombinations = allCombinations.OrderBy(x => random.Next()).Take(20).ToList();
                    Console.WriteLine($"  共生成 {allCombinations.Count} 个可能的组合，随机选择 {selectedCombinations.Count} 个进行测试");
                }
                else
                {
                    Console.WriteLine($"  共生成 {allCombinations.Count} 个可能的组合，全部进行测试");
                }
                
                int combinationIndex = 0;
                foreach (var (startIndex, length) in selectedCombinations)
                {
                    combinationIndex++;
                    searchSequenceNumber++;
                    System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] Seg {segIndex + 1}, 组合搜索 {combinationIndex}/{selectedCombinations.Count}: 起始位置={startIndex}, 长度={length}");
                    
                    // 获取组合中的句子名称
                    var combinationNames = sentenceNameList.Skip(startIndex).Take(length).ToList();
                    string combinationDisplay = string.Join("+", combinationNames);
                    System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] 组合句子名称: {combinationDisplay}");
                    
                    // 获取所有句子的内容
                    var contents = new List<string>();
                    bool allFound = true;
                    foreach (var name in combinationNames)
                    {
                        if (nameToContentMap.TryGetValue(name, out string content))
                        {
                            contents.Add(content);
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}]   句子 {name} 内容长度: {content.Length}");
                        }
                        else
                        {
                            allFound = false;
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}]   ❌ 在映射表中找不到句子名称: {name}");
                            break;
                        }
                    }
                    
                    if (allFound && contents.Count > 0)
                    {
                        // 拼接所有句子的内容
                        // 注意：映射表中的句子已经保留了末尾的 \r，所以直接拼接即可
                        string combinedContent = string.Join("", contents);
                        int totalLength = combinedContent.Length;
                        System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] 组合内容总长度: {totalLength}");
                        
                        // 注意：WordRangeFinder.FindSentenceRangeInDocument 现在支持超过250字符的文本
                        // 它会自动使用前250字符和后250字符进行包围搜索
                        Word.Range foundRange = WordRangeFinder.FindSentenceRangeInDocument(doc, combinedContent, combinationDisplay);
                        if (foundRange != null)
                        {
                            successCount++;
                            Console.WriteLine($"  [{searchSequenceNumber}] ✅ [{length}个句子] {combinationDisplay}: 搜索成功 (长度: {totalLength})");
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ✅ 搜索成功，Range位置: Start={foundRange.Start}, End={foundRange.End}");
                        }
                        else
                        {
                            failCount++;
                            string escapedContent = combinedContent.Replace("\r", "\\r").Replace("\n", "\\n");
                            Console.WriteLine($"  [{searchSequenceNumber}] ❌ [{length}个句子] {combinationDisplay}: 搜索失败 (长度: {totalLength})");
                            Console.WriteLine($"     完整搜索内容: {escapedContent}");
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ❌ 搜索失败：未在文档中找到匹配内容");
                            System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] 完整搜索内容: {escapedContent}");
                        }
                    }
                    else
                    {
                        failCount++;
                        var missingNames = combinationNames.Where(n => !nameToContentMap.ContainsKey(n)).ToList();
                        Console.WriteLine($"  [{searchSequenceNumber}] ❌ [{length}个句子] {combinationDisplay}: 在映射表中找不到 {string.Join(",", missingNames)} 的内容");
                        System.Diagnostics.Debug.WriteLine($"[测试] [{searchSequenceNumber}] ❌ 在映射表中找不到句子名称: {string.Join(",", missingNames)}");
                    }
                }
            }

            return (successCount, failCount);
        }

    }
}

