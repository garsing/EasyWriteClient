using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WordAddIn1.DocumentMapping;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// Word文档内容提取器
    /// 用于提取Word文档的完整内容，包括段落、表格、图片等
    /// </summary>
    public static class WordDocumentExtractor
    {
        /// <summary>
        /// readText 字符数超过此阈值时走后端 API（curr_doc_chunk_and_snapshot），否则本地 ProcessDocumentLocally。
        /// 可通过 config.json → App.BackendProcessTextLengthThreshold 调整。
        /// </summary>
        private static int BackendProcessTextLengthThreshold =>
            ConfigManager.Config?.App?.BackendProcessTextLengthThreshold ?? 60000;

        private static bool ShouldLogDocumentExtractDetails(bool optionsVerbose = false)
        {
            return optionsVerbose || EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentExtract);
        }

        private static void ExtractDetailLog(bool enabled, string message)
        {
            if (enabled)
            {
                System.Diagnostics.Debug.WriteLine(message);
            }
        }

        /// <summary>
        /// 使用正则表达式查找标签位置（支持带属性的标签）
        /// 例如：&lt;table&gt;、&lt;table id=T_eStyRRUe&gt;、&lt;cell colspan=3&gt; 等都能正确匹配
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="tagName">标签名称（如 "body", "table", "cell"）</param>
        /// <param name="isClosingTag">是否为闭合标签（如 &lt;/body&gt;）</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <returns>找到的标签位置，如果未找到返回 -1</returns>
        private static int FindTagPosition(string text, string tagName, bool isClosingTag, int startPos)
        {
            if (startPos >= text.Length)
            {
                return -1;
            }

            string pattern;
            if (isClosingTag)
            {
                // 闭合标签：</tagName>（不再包含换行符）
                // 例如：</cell>、</row>、</table>
                pattern = $@"</{tagName}\s*>";
            }
            else
            {
                // 开始标签：<tagName> 或 <tagName ...>（支持带属性的标签，如 <table id=T_eStyRRUe>）
                pattern = $@"<{tagName}(?:\s+[^>]*)?>";
            }

            Match match = Regex.Match(text.Substring(startPos), pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return startPos + match.Index;
            }

            return -1;
        }

        /// <summary>
        /// 查找 image 标签的位置（必须是 <image />\r，必须带回车符）
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <returns>找到的标签位置，如果未找到返回 -1</returns>
        private static int FindImageTagPosition(string text, int startPos)
        {
            if (startPos >= text.Length)
            {
                return -1;
            }

            // 匹配 <image />\r（必须带回车符，支持自闭合标签，可能有属性）
            string pattern = @"<image\s+[^>]*/>\r";
            Match match = Regex.Match(text.Substring(startPos), pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return startPos + match.Index;
            }

            return -1;
        }

        /// <summary>
        /// 查找 Chart 开始标签的位置（必须是 <Chart>）
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <returns>找到的标签位置，如果未找到返回 -1</returns>
        private static int FindChartTagPosition(string text, int startPos)
        {
            if (startPos >= text.Length)
            {
                return -1;
            }

            // 匹配 <Chart>（支持带属性的标签）
            string pattern = @"<Chart(?:\s+[^>]*)?>";
            Match match = Regex.Match(text.Substring(startPos), pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return startPos + match.Index;
            }

            return -1;
        }

        /// <summary>
        /// 查找 Chart 结束标签的位置（必须是 </Chart>\r，必须带回车符）
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <returns>找到的标签位置，如果未找到返回 -1</returns>
        private static int FindChartCloseTagPosition(string text, int startPos)
        {
            if (startPos >= text.Length)
            {
                return -1;
            }

            // 匹配 </Chart>\r（必须带回车符）
            string pattern = @"</Chart\s*>\r";
            Match match = Regex.Match(text.Substring(startPos), pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return startPos + match.Index;
            }

            return -1;
        }

        /// <summary>
        /// 查找第一个开始收集标志的位置和完整标签字符串
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <returns>找到的开始标志信息（位置、类型、完整标签字符串），如果未找到返回 null</returns>
        private static (int position, int type, string fullTag)? FindStartCollectFlag(string text, int startPos)
        {
            // 定义五个开始标志的正则表达式
            string bodyPattern = @"<body(?:\s+[^>]*)?>";
            string cellPattern = @"<cell(?:\s+[^>]*)?>";
            string tableClosePattern = @"</table\s*>";  // </table> 不再包含换行符
            string imagePattern = @"<image\s+[^>]*/>\r";  // 必须是 <image />\r，必须带回车符
            string chartClosePattern = @"</Chart\s*>\r";  // 必须是 </Chart>\r，必须带回车符

            int minPos = int.MaxValue;
            int type = -1;
            string fullTag = "";

            // 查找 <body>
            Match bodyMatch = Regex.Match(text.Substring(startPos), bodyPattern, RegexOptions.IgnoreCase);
            if (bodyMatch.Success)
            {
                int bodyPos = startPos + bodyMatch.Index;
                if (bodyPos >= startPos && bodyPos < minPos)
                {
                    minPos = bodyPos;
                    type = 1; // <body>
                    fullTag = bodyMatch.Value;
                }
            }

            // 查找 <cell>
            Match cellMatch = Regex.Match(text.Substring(startPos), cellPattern, RegexOptions.IgnoreCase);
            if (cellMatch.Success)
            {
                int cellPos = startPos + cellMatch.Index;
                if (cellPos >= startPos && cellPos < minPos)
                {
                    minPos = cellPos;
                    type = 2; // <cell>
                    fullTag = cellMatch.Value;
                }
            }

            // 查找 </table>（不再包含换行符）
            Match tableCloseMatch = Regex.Match(text.Substring(startPos), tableClosePattern, RegexOptions.IgnoreCase);
            if (tableCloseMatch.Success)
            {
                int tableClosePos = startPos + tableCloseMatch.Index;
                if (tableClosePos >= startPos && tableClosePos < minPos)
                {
                    minPos = tableClosePos;
                    type = 3; // </table>
                    fullTag = tableCloseMatch.Value;
                }
            }

            // 查找 <image />\r（必须是带回车符的）
            Match imageMatch = Regex.Match(text.Substring(startPos), imagePattern, RegexOptions.IgnoreCase);
            if (imageMatch.Success)
            {
                int imagePos = startPos + imageMatch.Index;
                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    type = 4; // <image />\r
                    fullTag = imageMatch.Value;
                }
            }

            // 查找 </Chart>\r（必须是带回车符的）
            Match chartCloseMatch = Regex.Match(text.Substring(startPos), chartClosePattern, RegexOptions.IgnoreCase);
            if (chartCloseMatch.Success)
            {
                int chartClosePos = startPos + chartCloseMatch.Index;
                if (chartClosePos >= startPos && chartClosePos < minPos)
                {
                    minPos = chartClosePos;
                    type = 5; // </Chart>\r
                    fullTag = chartCloseMatch.Value;
                }
            }

            if (minPos != int.MaxValue)
            {
                return (minPos, type, fullTag);
            }

            return null;
        }

        /// <summary>
        /// 查找对应的结束收集标志的位置和类型
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="startPos">开始搜索的位置</param>
        /// <param name="startType">开始标志类型（1: <body>, 2: <cell>, 3: </table>, 4: <image />, 5: </Chart>\r）</param>
        /// <returns>找到的结束标志位置和类型（1: <table>, 2: </cell>, 3: </body>, 4: <image />, 5: <Chart>），如果未找到返回 (-1, -1)</returns>
        private static (int position, int endType) FindEndCollectFlag(string text, int startPos, int startType)
        {
            int minPos = int.MaxValue;
            int endType = -1;

            if (startType == 1)
            {
                // 开始标志：<body>，结束标志：<table> 或 </body> 或 <image />\r 或 <Chart>
                int tablePos = FindTagPosition(text, "table", false, startPos);
                int bodyClosePos = FindTagPosition(text, "body", true, startPos);
                int imagePos = FindImageTagPosition(text, startPos);
                int chartPos = FindChartTagPosition(text, startPos);

                if (tablePos >= startPos && tablePos < minPos)
                {
                    minPos = tablePos;
                    endType = 1; // <table>
                }

                if (bodyClosePos >= startPos && bodyClosePos < minPos)
                {
                    minPos = bodyClosePos;
                    endType = 3; // </body>
                }

                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    endType = 4; // <image />\r
                }

                if (chartPos >= startPos && chartPos < minPos)
                {
                    minPos = chartPos;
                    endType = 5; // <Chart>
                }
            }
            else if (startType == 2)
            {
                // 开始标志：<cell>，结束标志：<table> 或 </cell> 或 <image />\r 或 <Chart>
                int tablePos = FindTagPosition(text, "table", false, startPos);
                int cellClosePos = FindTagPosition(text, "cell", true, startPos);
                int imagePos = FindImageTagPosition(text, startPos);
                int chartPos = FindChartTagPosition(text, startPos);

                if (tablePos >= startPos && tablePos < minPos)
                {
                    minPos = tablePos;
                    endType = 1; // <table>
                }

                if (cellClosePos >= startPos && cellClosePos < minPos)
                {
                    minPos = cellClosePos;
                    endType = 2; // </cell>
                }

                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    endType = 4; // <image />\r
                }

                if (chartPos >= startPos && chartPos < minPos)
                {
                    minPos = chartPos;
                    endType = 5; // <Chart>
                }
            }
            else if (startType == 3)
            {
                // 开始标志：</table>，结束标志：</cell> 或 </body> 或 <table> 或 <image />\r 或 <Chart>
                int cellClosePos = FindTagPosition(text, "cell", true, startPos);
                int bodyClosePos = FindTagPosition(text, "body", true, startPos);
                int tablePos = FindTagPosition(text, "table", false, startPos);
                int imagePos = FindImageTagPosition(text, startPos);
                int chartPos = FindChartTagPosition(text, startPos);

                if (cellClosePos >= startPos && cellClosePos < minPos)
                {
                    minPos = cellClosePos;
                    endType = 2; // </cell>
                }

                if (bodyClosePos >= startPos && bodyClosePos < minPos)
                {
                    minPos = bodyClosePos;
                    endType = 3; // </body>
                }

                if (tablePos >= startPos && tablePos < minPos)
                {
                    minPos = tablePos;
                    endType = 1; // <table>
                }

                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    endType = 4; // <image />\r
                }

                if (chartPos >= startPos && chartPos < minPos)
                {
                    minPos = chartPos;
                    endType = 5; // <Chart>
                }
            }
            else if (startType == 4)
            {
                // 开始标志：<image />\r，结束标志：<table> 或 </cell> 或 </body> 或 <image />\r 或 <Chart>
                int tablePos = FindTagPosition(text, "table", false, startPos);
                int cellClosePos = FindTagPosition(text, "cell", true, startPos);
                int bodyClosePos = FindTagPosition(text, "body", true, startPos);
                int imagePos = FindImageTagPosition(text, startPos);
                int chartPos = FindChartTagPosition(text, startPos);

                if (tablePos >= startPos && tablePos < minPos)
                {
                    minPos = tablePos;
                    endType = 1; // <table>
                }

                if (cellClosePos >= startPos && cellClosePos < minPos)
                {
                    minPos = cellClosePos;
                    endType = 2; // </cell>
                }

                if (bodyClosePos >= startPos && bodyClosePos < minPos)
                {
                    minPos = bodyClosePos;
                    endType = 3; // </body>
                }

                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    endType = 4; // <image />\r
                }

                if (chartPos >= startPos && chartPos < minPos)
                {
                    minPos = chartPos;
                    endType = 5; // <Chart>
                }
            }
            else if (startType == 5)
            {
                // 开始标志：</Chart>\r，结束标志：</cell> 或 </body> 或 <table> 或 <image />\r 或 <Chart>
                int cellClosePos = FindTagPosition(text, "cell", true, startPos);
                int bodyClosePos = FindTagPosition(text, "body", true, startPos);
                int tablePos = FindTagPosition(text, "table", false, startPos);
                int imagePos = FindImageTagPosition(text, startPos);
                int chartPos = FindChartTagPosition(text, startPos);

                if (cellClosePos >= startPos && cellClosePos < minPos)
                {
                    minPos = cellClosePos;
                    endType = 2; // </cell>
                }

                if (bodyClosePos >= startPos && bodyClosePos < minPos)
                {
                    minPos = bodyClosePos;
                    endType = 3; // </body>
                }

                if (tablePos >= startPos && tablePos < minPos)
                {
                    minPos = tablePos;
                    endType = 1; // <table>
                }

                if (imagePos >= startPos && imagePos < minPos)
                {
                    minPos = imagePos;
                    endType = 4; // <image />\r
                }

                if (chartPos >= startPos && chartPos < minPos)
                {
                    minPos = chartPos;
                    endType = 5; // <Chart>
                }
            }

            return minPos != int.MaxValue ? (minPos, endType) : (-1, -1);
        }

        /// <summary>
        /// 获取标签的结束位置（包括 > 符号）
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="tagStartPos">标签开始位置（< 的位置）</param>
        /// <returns>标签结束位置（> 的位置 + 1）</returns>
        private static int GetTagEndPosition(string text, int tagStartPos)
        {
            int tagEndPos = text.IndexOf('>', tagStartPos);
            if (tagEndPos >= 0)
            {
                return tagEndPos + 1;
            }
            return tagStartPos;
        }

        /// <summary>
        /// 使用状态机算法处理文本，按顺序提取 seg（只提取内容，不计算位置）
        /// </summary>
        /// <param name="text">输入文本</param>
        /// <param name="segs">seg 列表（输出参数，只存储内容）</param>
        /// <param name="segMarks">seg 标记列表（输出参数，1表示cell类型，0表示其他类型）</param>
        private static void ExtractSegsWithStateMachine(string text, List<string> segs, List<int> segMarks = null, bool detailLog = false)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 开始处理文本，长度: {text.Length}");
            ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 文本前100字符: {TextUtils.EscapeNewlinesForDb(text.Substring(0, Math.Min(100, text.Length)))}");

            int currentPos = 0;
            bool isCollecting = false; // false = 空置状态, true = 收集状态
            int startCollectPos = 0; // 开始收集的位置
            int currentStartFlagType = -1; // 当前开始标志类型（1=<body>, 2=<cell>, 3=</table>, 4=<image />\r, 5=</Chart>\r）

            while (currentPos < text.Length)
            {
                if (!isCollecting)
                {
                    // 空置状态：查找第一个开始收集标志
                    var startFlag = FindStartCollectFlag(text, currentPos);
                    if (!startFlag.HasValue)
                    {
                        // 没有找到开始收集标志，处理结束
                        ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 未找到更多开始收集标志，处理结束");
                        break;
                    }

                    int startFlagPos = startFlag.Value.position;
                    currentStartFlagType = startFlag.Value.type;
                    string fullTag = startFlag.Value.fullTag;

                    ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 找到开始收集标志，位置: {startFlagPos}, 类型: {currentStartFlagType} (1=<body>, 2=<cell>, 3=</table>, 4=<image />\r, 5=</Chart>\r), 完整标签: {fullTag}");

                    // 丢弃开始收集标志之前的内容（因为处于空置状态）
                    if (startFlagPos > currentPos)
                    {
                        string discarded = text.Substring(currentPos, startFlagPos - currentPos);
                        ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 丢弃空置状态下的内容，长度: {discarded.Length}");
                    }

                    // 第二部分：从开始标志位置开始的文本
                    string remainingText = text.Substring(startFlagPos);
                    
                    // 如果第二部分开头是完整标签字符串，减去它
                    if (remainingText.StartsWith(fullTag))
                    {
                        startCollectPos = startFlagPos + fullTag.Length; // 开始收集的位置是完整标签之后
                    }
                    else
                    {
                        // 如果格式不匹配，使用原来的方法作为后备
                        int startFlagEndPos = GetTagEndPosition(text, startFlagPos);
                        startCollectPos = startFlagEndPos;
                    }
                    
                    isCollecting = true; // 切换到收集状态
                    currentPos = startCollectPos;
                }
                else
                {
                    // 收集状态：查找对应的结束收集标志
                    var (endFlagPos, endFlagType) = FindEndCollectFlag(text, currentPos, currentStartFlagType);
                    string segContent;
                    
                    if (endFlagPos < 0)
                    {
                        // 没有找到结束标志，将剩余内容作为 seg
                        segContent = text.Substring(startCollectPos);
                        if (!string.IsNullOrEmpty(segContent))
                        {
                            segs.Add(segContent);
                            // 标记：只有当开始标志是 <cell>（type=2）且结束标志是 </cell>（endType=2）时才标记为 1
                            // 如果没有找到结束标志，即使开始标志是 <cell>，也标记为 0（因为不是完整的 cell）
                            int mark = 0;
                            if (segMarks != null)
                            {
                                segMarks.Add(mark);
                            }
                            ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] ✅ 添加 seg（未找到结束标志，剩余全部内容），开始类型: {currentStartFlagType}，标记: {mark}，当前 seg 数量: {segs.Count}");
                        }
                        break;
                    }

                    ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 找到结束收集标志，位置: {endFlagPos}，结束类型: {endFlagType} (1=<table>, 2=</cell>, 3=</body>, 4=<image />, 5=<Chart>)");

                    // 收集从 startCollectPos 到 endFlagPos 的内容作为 seg
                    segContent = text.Substring(startCollectPos, endFlagPos - startCollectPos);
                    if (!string.IsNullOrEmpty(segContent))
                    {
                        segs.Add(segContent);
                        // 标记：只有当开始标志是 <cell>（type=2）且结束标志是 </cell>（endType=2）且 seg 长度少于 150 字时才标记为 1
                        // 如果 <cell> 遇到其他结束标志（如 <table> 或 <image />），或长度 >= 150 字，标记为 0
                        int mark = (currentStartFlagType == 2 && endFlagType == 2 && segContent.Length < 150) ? 1 : 0;
                        if (segMarks != null)
                        {
                            segMarks.Add(mark);
                        }
                        string markReason = "";
                        if (currentStartFlagType == 2 && endFlagType == 2)
                        {
                            if (segContent.Length >= 150)
                            {
                                markReason = $"（长度 {segContent.Length} >= 150）";
                            }
                            else
                            {
                                markReason = "（cell类型且长度 < 150）";
                            }
                        }
                        else
                        {
                            markReason = "（非cell类型）";
                        }
                        ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] ✅ 添加 seg，位置: {startCollectPos}-{endFlagPos}，长度: {segContent.Length}，开始类型: {currentStartFlagType}，结束类型: {endFlagType}，标记: {mark}{markReason}，当前 seg 数量: {segs.Count}");
                    }
                    else
                    {
                        ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] ⚠️ seg 内容为空，不添加");
                    }

                    // 状态回到空置
                    isCollecting = false;
                    currentStartFlagType = -1;
                    currentPos = endFlagPos;
                }
            }

            ExtractDetailLog(detailLog, $"[ExtractSegsWithStateMachine] 处理完成，共提取 {segs.Count} 个 seg");
        }

        // 用于替换\r的特殊字符（Unicode私有使用区字符，XML不会处理）
        private const char CR_REPLACEMENT = '\uE000';

        /// <summary>
        /// 转义 HTML 实体，将特殊字符转换为 HTML 实体
        /// </summary>
        /// <param name="text">原始文本</param>
        /// <returns>转义后的文本</returns>
        private static string EscapeHtmlEntities(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // 注意：必须先转义 &，否则其他字符中的 & 会被误转义
            return text
                .Replace("&", "&amp;")
                .Replace("<", "&lt;")
                .Replace(">", "&gt;")
                .Replace("\"", "&quot;")
                .Replace("'", "&apos;");
        }

        /// <summary>
        /// 按照 cell 分割文本，返回 seg 和 seg 之间的内容（非 seg 内容）
        /// 使用状态机算法提取 seg
        /// </summary>
        /// <param name="text">输入文本（可能包含 &lt;table&gt; 标签）</param>
        /// <param name="segMarks">seg 标记列表（输出参数，1表示cell类型，0表示其他类型）</param>
        /// <returns>seg 和 seg 之间的内容列表，从非 seg 开始：非seg, seg, 非seg, seg, ...</returns>
        public static List<string> SplitTextByCellsForMapping(string text, out List<int> segMarks, bool detailLog = false)
        {
            var chunkParts = SplitTextByCells(text, out segMarks, detailLog);
            for (int i = 0; i < chunkParts.Count; i++)
            {
                chunkParts[i] = UnescapeHtmlEntities(chunkParts[i]);
            }

            return chunkParts;
        }

        private static List<string> SplitTextByCells(string text, out List<int> segMarks, bool detailLog = false)
        {
            segMarks = new List<int>();
            
            if (string.IsNullOrEmpty(text))
            {
                return new List<string>();
            }



            // 第一步：先将文本中的CR_REPLACEMENT转换回\r（如果之前处理过）
            text = text.Replace(CR_REPLACEMENT, '\r');

            // 第二步：用<body></body>包裹文本，以便状态机算法能够正确识别开始收集标志
            // 注意：状态机算法会从<body>标签之后开始收集，所以提取的seg不包含<body>和</body>标签
            string wrappedText = $"<body>{text}</body>";

            // 第三步：使用状态机算法按顺序提取所有 seg（只存储内容，不计算位置）
            var segContents = new List<string>();
            var tempSegMarks = new List<int>();
            ExtractSegsWithStateMachine(wrappedText, segContents, tempSegMarks, detailLog);
            
            // 保存 seg 标记
            segMarks = tempSegMarks;
            
            // 如果没有找到任何 seg，整个文本作为非 seg
            if (segContents.Count == 0)
            {
                return new List<string> { text };
            }
            
            // 第四步：按顺序"切蛋糕"的方式分割文本
            // 把所有文本拿出来，用 seg 按顺序切分，seg 之间的部分就是非 seg
            // 注意：无论 seg 在文本中出现多少次，都只取第一次出现的位置来切分
            var parts = new List<string>();
            string remainingText = text;  // 剩余待处理的文本
            
            foreach (string segContent in segContents)
            {
                // 在剩余文本中查找当前 seg（从开头开始，IndexOf 会返回第一次出现的位置）
                // 无论 seg 在 remainingText 中出现多少次，都只取第一次出现的位置
                int segPos = remainingText.IndexOf(segContent, 0);
                
                if (segPos < 0)
                {
                    // 如果找不到，记录调试信息
                    ExtractDetailLog(detailLog, $"[SplitTextByCells] ⚠️ 无法找到 seg（长度: {segContent.Length}）: {TextUtils.EscapeNewlinesForDb(segContent.Substring(0, Math.Min(50, segContent.Length)))}");
                    ExtractDetailLog(detailLog, $"[SplitTextByCells] 剩余文本（前100字符）: {TextUtils.EscapeNewlinesForDb(remainingText.Substring(0, Math.Min(100, remainingText.Length)))}");
                    continue;
                }
                
                // seg 之前的部分作为非 seg（使用第一次出现的位置）
                if (segPos > 0)
                {
                    string nonSeg = remainingText.Substring(0, segPos);
                    parts.Add(nonSeg);  // 添加非 seg
                }
                else
                {
                    // 如果 segPos == 0，说明剩余文本一开始就是 seg，添加空字符串作为非 seg
                    parts.Add("");
                }

                // 添加 seg 内容
                parts.Add(segContent);

                // 更新剩余文本为 seg 之后的部分（从第一次出现的位置之后开始）
                remainingText = remainingText.Substring(segPos + segContent.Length);
            }

            // 添加最后一个 seg 之后的内容（非 seg）
            if (!string.IsNullOrEmpty(remainingText))
            {
                parts.Add(remainingText);
            }
            else if (parts.Count > 0 && parts.Count % 2 == 0)
            {
                // 如果剩余文本为空但前面有 seg，需要添加空字符串作为非 seg
                parts.Add("");
            }

            // 第五步：确保所有parts中的CR_REPLACEMENT都转换回\r
            for (int i = 0; i < parts.Count; i++)
            {
                if (parts[i].Contains(CR_REPLACEMENT))
                {
                    parts[i] = parts[i].Replace(CR_REPLACEMENT, '\r');
                }
            }

            // 【调试输出】打印所有 seg 和非 seg（全量，不省略）
            if (detailLog)
            {
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine("=== SplitTextByCells 分割结果（全量）===");
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine($"总共 {parts.Count} 个部分");
                for (int i = 0; i < parts.Count; i++)
                {
                    string part = parts[i];
                    string partType = (i % 2 == 0) ? "非seg" : "seg";
                    string escapedPart = TextUtils.EscapeNewlinesForDb(part);
                    System.Diagnostics.Debug.WriteLine($"--- 部分 {i + 1} ({partType}) ---");
                    System.Diagnostics.Debug.WriteLine($"长度: {part.Length} 字符");
                    // 如果是 seg，打印标记位
                    if (i % 2 == 1)
                    {
                        int segIndex = (i - 1) / 2;
                        if (segIndex < segMarks.Count)
                        {
                            int mark = segMarks[segIndex];
                            string markType = (mark == 1) ? "cell" : "其他";
                            System.Diagnostics.Debug.WriteLine($"标记: {mark} ({markType})");
                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"标记: (未找到，索引 {segIndex} 超出范围)");
                        }
                    }
                    System.Diagnostics.Debug.WriteLine($"内容（已转义，全量）: {escapedPart}");
                    System.Diagnostics.Debug.WriteLine("");
                }
                System.Diagnostics.Debug.WriteLine("==========================================");
            }

            return parts;
        }

        /// <summary>
        /// 切分文本为句子列表（使用 SentenceSplitter）
        /// </summary>
        /// <param name="text">要切分的文本</param>
        /// <returns>句子列表</returns>
        private static List<string> SplitIntoSentences(string text)
        {
            return SentenceSplitter.SplitForReadback(text);
        }

        /// <summary>
        /// 第四步：文本分块
        /// 对处理后的文本进行分块，目前只分一块（返回整个文本作为一个chunk）
        /// </summary>
        /// <param name="readText">处理后的文本（带标签）</param>
        /// <returns>带标签的分块列表</returns>
        private static List<string> ChunkText(string readText)
        {
            if (string.IsNullOrEmpty(readText))
            {
                return new List<string>();
            }

            // 目前只分一块，返回整个文本作为一个chunk
            return new List<string> { readText };
        }

        /// <summary>
        /// 处理文档的完整流程（从文件读入到生成快照）
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <returns>处理结果，包含快照和映射信息</returns>
        public static ProcessingResult ProcessDocument(Word.Document document)
        {
            return ProcessDocument(document, new ProcessDocumentOptions());
        }

        /// <param name="extractFormat">是否运行 FormatExtractor 并写入 DocumentState.SubtypeFormatMapping</param>
        public static ProcessingResult ProcessDocument(Word.Document document, bool extractFormat)
        {
            return ProcessDocument(document, new ProcessDocumentOptions { ExtractFormat = extractFormat });
        }

        /// <param name="extractFormat">是否运行 FormatExtractor 并写入 DocumentState.SubtypeFormatMapping</param>
        /// <param name="snapshotSource">写入快照历史时的来源标记</param>
        public static ProcessingResult ProcessDocument(Word.Document document, bool extractFormat, string snapshotSource)
        {
            return ProcessDocument(document, new ProcessDocumentOptions
            {
                ExtractFormat = extractFormat,
                SnapshotSource = snapshotSource
            });
        }

        public static ProcessingResult ProcessDocument(Word.Document document, ProcessDocumentOptions options)
        {
            if (options == null)
            {
                options = new ProcessDocumentOptions();
            }

            bool verbose = options.VerboseDebug;
            bool detailLog = ShouldLogDocumentExtractDetails(verbose);
            ExtractFormatMode extractFormatMode = options.ExtractFormatMode;
            string snapshotSource = options.SnapshotSource;

            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"{FormatInheritHelper.DbgPrefix} ProcessDocument start extractFormatMode={extractFormatMode} extractImages={options.ExtractImages}");
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine("=== 开始处理文档 ===");
                System.Diagnostics.Debug.WriteLine("==========================================");
            }
            else
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ProcessDocument] 轻量模式 source={snapshotSource ?? "(null)"} " +
                    $"extractFormatMode={extractFormatMode} extractImages={options.ExtractImages} verbose=false");
            }

            // 第一步：绑定活动文档并确保 doc_uuid
            DocumentState.BindAndActivate(document);
            string docUuid = DocumentState.EnsureCurrDocUuid(document);
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine($"[第一步] 当前文档UUID: {docUuid}");
            }

            ProcessDocumentCache.EnsureFingerprintVersion();

            BodyFingerprintDetail bodyDetail = null;
            var bodyCheckStopwatch = System.Diagnostics.Stopwatch.StartNew();
            if (options.ForceRefresh)
            {
                bodyDetail = BodyXmlFingerprint.ComputeDetailed(document);
                bodyCheckStopwatch.Stop();
                ProcessDocumentCache.LogMiss(
                    snapshotSource,
                    bodyDetail.Fingerprint,
                    "force_refresh",
                    bodyCheckStopwatch.ElapsedMilliseconds);
            }
            else
            {
                bodyDetail = BodyXmlFingerprint.ComputeDetailed(document);
                bodyCheckStopwatch.Stop();
                if (ProcessDocumentCache.TryGetHit(
                    docUuid,
                    bodyDetail,
                    options,
                    bodyCheckStopwatch.ElapsedMilliseconds,
                    out ProcessingResult hitResult))
                {
                    // 缓存命中仍可能缺 BodyFormat（易失态被清）：BodyOnly 时补抽
                    if (extractFormatMode == ExtractFormatMode.BodyOnly && !DocumentState.HasBodyFormat)
                    {
                        try
                        {
                            var displays = hitResult.ChunkInfo != null
                                ? hitResult.ChunkInfo.Select(c => c.DisplayContent ?? "").ToList()
                                : new List<string>();
                            var nameMap = hitResult.NameToContentMap ?? new Dictionary<string, string>();
                            System.Diagnostics.Debug.WriteLine(
                                "[ProcessDocument] cache hit 但 BodyFormat 空 → 补跑 BodyFormatExtractor");
                            BodyFormatExtractor.ExtractAndSave(document, displays, nameMap);
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"[ProcessDocument] cache hit BodyOnly 补抽失败: {ex.Message}");
                        }
                    }

                    return hitResult;
                }
            }

            // 第一步半：ExtractImages 时 SaveAs+_copy 供解压抽图；轻量 API 见字数分支内 SaveDocumentCopy
            string docCopyPath = null;
            string tempFilePath = null;
            if (options.ExtractImages)
            {
                TrySaveDocumentCopies(document, createSeparateCopy: true, verbose, out docCopyPath, out tempFilePath);
            }

            // 第二步：文件读入
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine("📄 第二步：文件读入...");
            }

            string readText = WordReader.ReadWord(document, docCopyPath, options.ExtractImages);
            ExtractDetailLog(detailLog, $"✅ 文件读入完成，文本长度: {readText.Length} 字符");

            // 第二步半：移除 </cell>、</row> 和 </table> 后面的换行符 \n
            ExtractDetailLog(detailLog, "🔧 第二步半：移除表格与图表标签后的换行符...");
            readText = Regex.Replace(readText, @"</cell>\n", "</cell>");
            readText = Regex.Replace(readText, @"</row>\n", "</row>");
            readText = Regex.Replace(readText, @"</table>\n", "</table>");
            readText = Regex.Replace(readText, @"</Chart>\n", "</Chart>");
            readText = Regex.Replace(readText, @"<Chart>\n", "<Chart>");
            readText = Regex.Replace(readText, @"</DataTable>\n", "</DataTable>");
            readText = Regex.Replace(readText, @"<DataTable>\n", "<DataTable>");
            readText = Regex.Replace(readText, @"</Header>\n", "</Header>");
            readText = Regex.Replace(readText, @"<Header>\n", "<Header>");
            readText = Regex.Replace(readText, @"</Rows>\n", "</Rows>");
            readText = Regex.Replace(readText, @"<Rows>\n", "<Rows>");
            readText = Regex.Replace(readText, @"</Column>\n", "</Column>");
            readText = Regex.Replace(readText, @"<Column[^>]*>\n", m => m.Value.TrimEnd('\n'));
            readText = Regex.Replace(readText, @"</Row>\n", "</Row>");
            readText = Regex.Replace(readText, @"<Row>\n", "<Row>");

            
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine($"✅ 文件读入完成，文本: {TextUtils.EscapeNewlinesForDb(readText)}");
            }

            ExtractDetailLog(detailLog, $"✅ 移除换行符完成，文本长度: {readText.Length} 字符");

            // 第二步半：根据字数选择处理方式
            ProcessingResult result;
            if (readText.Length > BackendProcessTextLengthThreshold)
            {
                ExtractDetailLog(detailLog,
                    $"📊 文档字数 ({readText.Length}) 超过 {BackendProcessTextLengthThreshold}，使用API处理");
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"📊 文档字数 ({readText.Length}) 超过 {BackendProcessTextLengthThreshold}，使用API处理");
                }

                if (string.IsNullOrEmpty(docCopyPath))
                {
                    if (TrySaveDocumentCopyForApi(document, verbose, out docCopyPath, out tempFilePath))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[ProcessDocument] 轻量模式为 API 上传 SaveDocumentCopy: {docCopyPath}");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "[ProcessDocument] ⚠️ API 分支 SaveDocumentCopy 失败，上传将失败并可能回退 Local");
                    }
                }

                result = ProcessDocumentWithApi(document, docUuid, docCopyPath, tempFilePath, snapshotSource);
            }
            else
            {
                ExtractDetailLog(detailLog,
                    $"📊 文档字数 ({readText.Length}) 未超过 {BackendProcessTextLengthThreshold}，使用本地处理");
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"📊 文档字数 ({readText.Length}) 未超过 {BackendProcessTextLengthThreshold}，使用本地处理");
                }

                result = ProcessDocumentLocally(readText, docCopyPath, tempFilePath, options);
            }
            
            // 第三步：格式抽取 — None 跳过；BodyOnly 轻量正文；Full 全量 subtype
            if (extractFormatMode == ExtractFormatMode.None)
            {
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"{FormatInheritHelper.DbgPrefix} ProcessDocument 跳过格式抽取（ExtractFormatMode=None）");
                }

                ProcessDocumentCache.SaveAfterFullRead(docUuid, bodyDetail, result, snapshotSource);
                return result;
            }

            // 2026-08：get 已改为 None；BodyOnly 分支保留供调试/回滚，主路径勿再走
            if (extractFormatMode == ExtractFormatMode.BodyOnly)
            {
                try
                {
                    System.Diagnostics.Debug.WriteLine("[格式提取] ExtractFormatMode=BodyOnly → BodyFormatExtractor");
                    var displayContents = new List<string>();
                    if (options.BuildDisplayContent && result.ChunkInfo != null)
                    {
                        displayContents = result.ChunkInfo
                            .Select(c => c.DisplayContent ?? "")
                            .ToList();
                    }

                    var nameMap = result.NameToContentMap ?? new Dictionary<string, string>();
                    BodyFormatExtractor.ExtractAndSave(document, displayContents, nameMap);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[格式提取] ⚠️ BodyOnly 异常: {ex.Message}");
                }

                ProcessDocumentCache.SaveAfterFullRead(docUuid, bodyDetail, result, snapshotSource);
                return result;
            }

            try
            {
                System.Diagnostics.Debug.WriteLine("");
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine("=== 开始提取格式（Full）===");
                System.Diagnostics.Debug.WriteLine("==========================================");
                
                System.Diagnostics.Debug.WriteLine("[格式提取] 使用 ProcessingResult.ReadText 按行分类并提取格式...");
                string fullText = result.ReadText;
                
                if (string.IsNullOrEmpty(readText))
                {
                    System.Diagnostics.Debug.WriteLine("[格式提取] ⚠️ ReadText 为空，跳过格式提取");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[格式提取] 调用 FormatExtractor.ExtractSubtypeFormatMapping...");
                    var formatExtractionStartTime = System.Diagnostics.Stopwatch.StartNew();

                    List<SampleSentence> sentencePool = null;
                    if (options.BuildDisplayContent && result.ChunkInfo != null)
                    {
                        var displayContents = result.ChunkInfo
                            .Select(c => c.DisplayContent ?? "")
                            .ToList();
                        var nameMap = result.NameToContentMap ?? new Dictionary<string, string>();
                        sentencePool = SamplingSentencePool.BuildSamplingSentencePool(displayContents, nameMap);
                        System.Diagnostics.Debug.WriteLine(
                            $"[格式提取] display→pool 句数: {sentencePool.Count}（BuildDisplayContent=true）");
                    }

                    var (classification, subtypeFormatMapping) =
                        FormatExtractor.ExtractSubtypeFormatMapping(readText, document, sentencePool);
                    
                    // Full 路径禁止同步 BodyFormat（D10）
                    FormatExtractor.SaveToDocumentState(subtypeFormatMapping, classification);
                    
                    formatExtractionStartTime.Stop();
                    
                    int headingLines = classification.HeadingsByLanguage.Values.Sum(l => l.Count);
                    System.Diagnostics.Debug.WriteLine($"[格式提取] ✅ 标题行 {headingLines}，正文候选行 {classification.BodyLines.Count}，耗时: {formatExtractionStartTime.ElapsedMilliseconds}ms ({formatExtractionStartTime.ElapsedMilliseconds / 1000.0:F2}秒)");
                    
                    System.Diagnostics.Debug.WriteLine("[格式提取] 按行输出标题分类摘要...");
                    var outputStartTime = System.Diagnostics.Stopwatch.StartNew();
                    
                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("按行分类摘要（标题按语言分组）");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("按行分类摘要（标题按语言分组）");
                    Console.WriteLine("==========================================");
                    
                    foreach (var langKv in classification.HeadingsByLanguage)
                    {
                        string langHeader = $"语言: {langKv.Key}，标题数: {langKv.Value.Count}";
                        System.Diagnostics.Debug.WriteLine(langHeader);
                        Console.WriteLine(langHeader);
                        foreach (Dictionary<string, object> h in langKv.Value)
                        {
                            string lineNum = h.ContainsKey("line") ? h["line"]?.ToString() ?? "" : "";
                            string dst = h.ContainsKey("detailed_subtype") ? h["detailed_subtype"]?.ToString() ?? "" : "";
                            string t = h.ContainsKey("text") ? h["text"]?.ToString() ?? "" : "";
                            string lineOut = $"  line {lineNum} [{dst}] {t}";
                            System.Diagnostics.Debug.WriteLine(lineOut);
                            Console.WriteLine(lineOut);
                        }
                    }
                    
                    string bodySummary = $"正文候选行数: {classification.BodyLines.Count}";
                    System.Diagnostics.Debug.WriteLine(bodySummary);
                    Console.WriteLine(bodySummary);
                    
                    outputStartTime.Stop();
                    System.Diagnostics.Debug.WriteLine($"[格式提取] ✅ 摘要输出耗时: {outputStartTime.ElapsedMilliseconds}ms");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("");
                    
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");
                    
                    // 输出JSON格式的映射关系
                    System.Diagnostics.Debug.WriteLine("");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("detailed_subtype 格式映射关系（JSON）");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    
                    string jsonOutput = JsonConvert.SerializeObject(subtypeFormatMapping, Formatting.Indented);
                    
                    System.Diagnostics.Debug.WriteLine(jsonOutput);
                    Console.WriteLine("");
                    Console.WriteLine("==========================================");
                    Console.WriteLine("detailed_subtype 格式映射关系（JSON）");
                    Console.WriteLine("==========================================");
                    Console.WriteLine(jsonOutput);
                    Console.WriteLine("==========================================");
                    Console.WriteLine("");
                    
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("=== 格式提取完成 ===");
                    System.Diagnostics.Debug.WriteLine("==========================================");
                    System.Diagnostics.Debug.WriteLine("");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[格式提取] ⚠️ 格式提取过程中发生异常: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[格式提取] ⚠️ 堆栈跟踪: {ex.StackTrace}");
                // 格式提取失败不影响主流程，继续执行
            }

            ProcessDocumentCache.SaveAfterFullRead(docUuid, bodyDetail, result, snapshotSource);
            return result;
        }

        /// <summary>
        /// 使用API处理文档（字数超过 <see cref="BackendProcessTextLengthThreshold"/> 时）
        /// </summary>
        /// <param name="document">Word文档对象</param>
        /// <param name="docUuid">文档UUID</param>
        /// <param name="docCopyPath">文档副本路径</param>
        /// <param name="tempFilePath">临时文件路径</param>
        /// <returns>处理结果</returns>
        private static ProcessingResult ProcessDocumentWithApi(Word.Document document, string docUuid, string docCopyPath, string tempFilePath, string snapshotSource)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[API] 文档UUID: {docUuid}");
                System.Diagnostics.Debug.WriteLine($"[API] 文件名: {Path.GetFileName(docCopyPath)}");

                // 仅在网络线程等待 HTTP；解析与 DocumentState/Word COM 更新留在当前 STA 线程，避免 .Result 死锁
                var httpResult = BackendApiClient.PostProcessDocumentMultipartAsync(docCopyPath, docUuid)
                    .ConfigureAwait(false)
                    .GetAwaiter()
                    .GetResult();
                var processingResult = ApplyProcessDocumentApiResponse(document, docUuid, snapshotSource, httpResult);
                if (processingResult != null)
                {
                    System.Diagnostics.Debug.WriteLine("✅ API处理文档成功，使用API返回的数据");
                    
                    // 不再清理临时文件，避免文件被占用导致的删除失败
                    // CleanupTempFiles(docCopyPath, tempFilePath);
                    
                    return processingResult;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("⚠️ API处理文档失败，回退到本地处理");
                    // API失败时回退到本地处理
                    // 需要重新读取文档内容
                    string readText = WordReader.ReadWord(document, docCopyPath);
                    return ProcessDocumentLocally(readText, docCopyPath, tempFilePath, new ProcessDocumentOptions { SnapshotSource = snapshotSource });
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ API处理文档异常: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"⚠️ 堆栈跟踪: {ex.StackTrace}");
                System.Diagnostics.Debug.WriteLine("⚠️ 回退到本地处理");
                
                // API异常时回退到本地处理
                string readText = WordReader.ReadWord(document, docCopyPath);
                return ProcessDocumentLocally(readText, docCopyPath, tempFilePath, new ProcessDocumentOptions { SnapshotSource = snapshotSource });
            }
        }

        /// <summary>
        /// 反转义 HTML 实体，将转义的字符还原为原始字符
        /// </summary>
        /// <param name="text">包含 HTML 实体的文本</param>
        /// <returns>反转义后的文本</returns>
        private static string UnescapeHtmlEntities(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            // 注意：必须先反转义其他实体，最后反转义 &amp;，否则会误替换其他实体中的 &
            return text
                .Replace("&lt;", "<")
                .Replace("&gt;", ">")
                .Replace("&quot;", "\"")
                .Replace("&apos;", "'")
                .Replace("&amp;", "&");
        }

        /// <summary>
        /// 使用本地方式处理文档（字数不超过 <see cref="BackendProcessTextLengthThreshold"/> 时）
        /// </summary>
        /// <param name="readText">处理后的文本</param>
        /// <param name="docCopyPath">文档副本路径</param>
        /// <param name="tempFilePath">临时文件路径</param>
        /// <returns>处理结果</returns>
        private static ProcessingResult ProcessDocumentLocally(
            string readText,
            string docCopyPath,
            string tempFilePath,
            ProcessDocumentOptions options)
        {
            if (options == null)
            {
                options = new ProcessDocumentOptions();
            }

            bool verbose = options.VerboseDebug;
            bool detailLog = ShouldLogDocumentExtractDetails(verbose);
            string snapshotSource = options.SnapshotSource;

            // 第四步：分块（只分一块）
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine("📦 第四步：文本分块...");
            }
            //readText中正文的内容已经转义，但是表格的内容是正常的标签。
            var chunks = ChunkText(readText);
            ExtractDetailLog(detailLog, $"✅ 文本分块完成，共 {chunks.Count} 个chunk");

            // 第六步：分句和映射（MapReadTextToCodes，S_ + P_ 同次）
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine("📋 第六步：分句和映射（MapReadTextToCodes）...");
            }

            MapReadTextResult mapResult = MapReadTextToCodes.MapChunks(chunks);
            DocumentState.SetChunkParagraphDisplayContents(mapResult.ChunkParagraphDisplayContents);

            var nameToContentMap = mapResult.NameToSentenceContentMap ?? new Dictionary<string, string>();
            var allChunkSentenceNames = mapResult.AllChunkSentenceNames;
            var allChunkSegsSentenceNames = mapResult.AllChunkSegsSentenceNames;
            var allChunkNonSegs = mapResult.AllChunkNonSegs;

            if (detailLog)
            {
                ExtractDetailLog(detailLog,
                    $"✅ MapReadTextToCodes 完成：S_={mapResult.SentenceMappings?.Count ?? 0} " +
                    $"P_={mapResult.ParagraphMappings?.Count ?? 0} " +
                    $"索引={mapResult.SentenceToParagraph?.Count ?? 0}");
            }

            ExtractDetailLog(detailLog, $"✅ 分句和映射完成，共处理 {chunks.Count} 个 chunk");

            // 第七步：display_content（已由 MapReadTextToCodes 构建 sentence display）
            var displayContents = new List<string>();
            if (options.BuildDisplayContent)
            {
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine("📝 第七步：使用 MapReadTextToCodes 的 display_content...");
                }

                displayContents.AddRange(mapResult.ChunkDisplayContents);
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine($"✅ display_content 就绪，共 {displayContents.Count} 个 chunk");
                }
            }
            else
            {
                for (int i = 0; i < chunks.Count; i++)
                {
                    displayContents.Add(string.Empty);
                }
            }

            // 第八步：生成快照
            if (verbose)
            {
                System.Diagnostics.Debug.WriteLine("📸 第八步：生成快照...");
            }

            string snapshot = BuildFlatSnapshot(allChunkSentenceNames);
            string segSnapshot;
            string segSnapshotPath;

            if (chunks.Count == 1)
            {
                segSnapshot = BuildSegSnapshotFromChunkSegs(allChunkSegsSentenceNames[0]);
                segSnapshotPath = "step6_fast";
            }
            else
            {
                var snapshotResult = GenerateSnapshots(
                    readText,
                    allChunkSentenceNames,
                    nameToContentMap,
                    displayContents,
                    verbose: verbose,
                    detailLog: detailLog);
                snapshot = snapshotResult.Snapshot;
                segSnapshot = snapshotResult.SegSnapshot;
                segSnapshotPath = "step8_full";
            }

#if DEBUG
            if (detailLog && chunks.Count == 1)
            {
                var legacyResult = GenerateSnapshots(
                    readText,
                    allChunkSentenceNames,
                    nameToContentMap,
                    displayContents,
                    verbose: true,
                    detailLog: true);
                if (!string.Equals(segSnapshot, legacyResult.SegSnapshot, StringComparison.Ordinal))
                {
                    System.Diagnostics.Debug.WriteLine("[seg_snapshot] WARN step6_fast != step8_full");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[seg_snapshot] OK step6_fast verified against step8_full");
                }
            }
#endif

            ExtractDetailLog(detailLog,
                $"[seg_snapshot] path={segSnapshotPath} chunks={chunks.Count} len={segSnapshot.Length}");

            if (verbose)
            {
                string escapedSegSnapshot = TextUtils.EscapeNewlinesForDb(segSnapshot);
                Console.WriteLine($"");
                Console.WriteLine($"==========================================");
                Console.WriteLine($"Seg Snapshot:");
                Console.WriteLine($"==========================================");
                Console.WriteLine(escapedSegSnapshot);
                Console.WriteLine($"==========================================");
                System.Diagnostics.Debug.WriteLine($"Seg Snapshot: {escapedSegSnapshot}");
                System.Diagnostics.Debug.WriteLine("==========================================");
                System.Diagnostics.Debug.WriteLine("=== 文档处理完成 ===");
                System.Diagnostics.Debug.WriteLine("==========================================");
            }

            // 同步快照到 DocumentState
            string historySource = string.IsNullOrEmpty(snapshotSource) ? "process_document_local" : snapshotSource;
            DocumentState.SetSnapshots(snapshot, segSnapshot, historySource);
            ExtractDetailLog(detailLog, "[同步] 快照已同步到 DocumentState");

            // 构建 ChunkInfo 列表
            var chunkInfoList = new List<ChunkInfo>();
            for (int i = 0; i < chunks.Count; i++)
            {
                var chunkInfo = new ChunkInfo
                {
                    ChunkIdx = i + 1,  // 从1开始
                    Content = chunks[i],
                    DisplayContent = displayContents[i],
                    Summary = ""  // 前端没有提取，为空字符串
                };
                chunkInfoList.Add(chunkInfo);

                if (verbose)
                {
                    Console.WriteLine($"");
                    Console.WriteLine($"==========================================");
                    Console.WriteLine($"📦 ChunkInfo (Chunk {chunkInfo.ChunkIdx}) - 构建完成");
                    Console.WriteLine($"==========================================");
                    Console.WriteLine($"Content 长度: {chunkInfo.Content.Length} 字符");
                    Console.WriteLine($"DisplayContent 长度: {chunkInfo.DisplayContent.Length} 字符");
                    Console.WriteLine($"Summary: \"{chunkInfo.Summary}\"");
                    string escapedDisplayContentForChunkInfo = TextUtils.EscapeNewlinesForDb(chunkInfo.DisplayContent);
                    Console.WriteLine($"DisplayContent (已转义，完整内容):");
                    Console.WriteLine(escapedDisplayContentForChunkInfo);
                    Console.WriteLine($"==========================================");
                    System.Diagnostics.Debug.WriteLine($"📦 ChunkInfo (Chunk {chunkInfo.ChunkIdx}) 构建完成");
                    System.Diagnostics.Debug.WriteLine($"  Content 长度: {chunkInfo.Content.Length} 字符");
                    System.Diagnostics.Debug.WriteLine($"  DisplayContent 长度: {chunkInfo.DisplayContent.Length} 字符");
                    System.Diagnostics.Debug.WriteLine($"  DisplayContent (已转义，完整): {escapedDisplayContentForChunkInfo}");
                }
            }

            // 生成表格序号到编号的映射
            var tableIndexToIdMap = new Dictionary<int, string>();
            var tableIdOrder = DocumentState.GetTableIdOrder();
            for (int i = 0; i < tableIdOrder.Count; i++)
            {
                tableIndexToIdMap[i + 1] = tableIdOrder[i]; // 序号从1开始
            }
            ExtractDetailLog(detailLog, $"✅ 生成表格序号到编号映射，共 {tableIndexToIdMap.Count} 个表格");

            // 不再清理临时文件，避免文件被占用导致的删除失败
            // CleanupTempFiles(docCopyPath, tempFilePath);

            // 对 readText 进行反转义
            readText = UnescapeHtmlEntities(readText);

            var imageIndexToIdMap = new Dictionary<int, string>();
            var imageIdOrder = DocumentState.GetImageIdOrder();
            for (int i = 0; i < imageIdOrder.Count; i++)
            {
                imageIndexToIdMap[i] = imageIdOrder[i];
            }

            return new ProcessingResult
            {
                ReadText = readText,
                ChunkInfo = chunkInfoList,
                NameToContentMap = nameToContentMap,
                Snapshot = snapshot,
                SegSnapshot = segSnapshot,
                TableIndexToIdMap = tableIndexToIdMap,
                ImageIndexToIdMap = imageIndexToIdMap,
                ProcessingMethod = "Local"
            };
        }
        

        /// <summary>
        /// 生成内容的SHA256哈希码（用于验证内容唯一性）
        /// 注意：必须与 DocumentState.GenerateContentHash 使用相同的格式（Base64），以确保hash一致性
        /// </summary>
        /// <param name="content">句子内容（包括空字符串）</param>
        /// <returns>Base64格式的哈希码</returns>
        private static string GenerateContentHash(string content)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content ?? ""));
                // 使用 Base64 格式，与 DocumentState.GenerateContentHash 保持一致
                return Convert.ToBase64String(hashBytes);
            }
        }

        private static string BuildFlatSnapshot(List<List<string>> allChunkSentenceNames)
        {
            var allSentenceNames = new List<string>();
            foreach (var chunkSentenceNames in allChunkSentenceNames)
            {
                allSentenceNames.AddRange(chunkSentenceNames);
            }

            return string.Join(",", allSentenceNames);
        }

        private static string BuildSegSnapshotFromChunkSegs(List<List<string>> chunkSegsSentenceNames)
        {
            if (chunkSegsSentenceNames == null || chunkSegsSentenceNames.Count == 0)
            {
                return string.Empty;
            }

            var chunkSegSnapshotParts = new List<string>();
            foreach (var segSentenceNames in chunkSegsSentenceNames)
            {
                chunkSegSnapshotParts.Add(string.Join(",", segSentenceNames));
            }

            return string.Join(";", chunkSegSnapshotParts);
        }

        /// <summary>
        /// 第八步：生成快照
        /// 生成snapshot（扁平化的句子名称列表）和seg_snapshot（分段的句子名称列表）
        /// 将processed_text当做只有一个分块，进行seg分割，查找名字和句子内容对应表，得到每个seg的名字，拼接成seg_snapshot
        /// </summary>
        /// <param name="readText">处理后的文本（带标签）</param>
        /// <param name="allChunkSentenceNames">所有chunk的句子名称列表（二维数组）</param>
        /// <param name="nameToContentMap">名字到句子内容的映射字典</param>
        /// <param name="displayContents">display_content列表（用于打印）</param>
        /// <returns>快照结果，包含snapshot和seg_snapshot</returns>
        private static (string Snapshot, string SegSnapshot) GenerateSnapshots(
            string readText,
            List<List<string>> allChunkSentenceNames,
            Dictionary<string, string> nameToContentMap,
            List<string> displayContents,
            bool verbose = false,
            bool detailLog = false)
        {
            // 将所有 chunk 的句子名称扁平化
            var allSentenceNames = new List<string>();
            foreach (var chunkSentenceNames in allChunkSentenceNames)
            {
                allSentenceNames.AddRange(chunkSentenceNames);
            }
            string snapshotStr = string.Join(",", allSentenceNames);
            string escapedSnapshotStr = TextUtils.EscapeNewlinesForDb(snapshotStr);
            ExtractDetailLog(detailLog, $"   原快照: {escapedSnapshotStr}");

            // 构建hash到name的映射字典，用于通过句子内容查找名字
            var hashToNameMap = new Dictionary<string, string>();
            foreach (var kvp in nameToContentMap)
            {
                string sentenceName = kvp.Key;
                string sentenceContent = kvp.Value;
                string contentHash = GenerateContentHash(sentenceContent);
                if (!string.IsNullOrEmpty(contentHash))
                {
                    hashToNameMap[contentHash] = sentenceName;
                }
            }
            ExtractDetailLog(detailLog, $"   构建hash到name映射完成，共 {hashToNameMap.Count} 个映射");

            // 对 processed_text 进行 seg 分割
            var chunkParts = SplitTextByCells(readText, out _, detailLog);
            
            // 对 chunkParts 进行反转义（因为 SplitTextByCells 内部已经处理完标签）
            for (int i = 0; i < chunkParts.Count; i++)
            {
                chunkParts[i] = UnescapeHtmlEntities(chunkParts[i]);
            }
            
            ExtractDetailLog(detailLog, $"   seg分割完成，共 {chunkParts.Count} 个部分");

            // 分离出 chunk_non_segs 和 chunk_segs
            var chunkNonSegs = new List<string>();
            var chunkSegs = new List<string>();

            for (int partIdx = 0; partIdx < chunkParts.Count; partIdx++)
            {
                if (partIdx % 2 == 0)  // 偶数索引（0, 2, 4...）是非 seg 内容
                {
                    chunkNonSegs.Add(chunkParts[partIdx]);
                }
                else  // 奇数索引（1, 3, 5...）是 seg
                {
                    chunkSegs.Add(chunkParts[partIdx]);
                }
            }
            ExtractDetailLog(detailLog, $"   分离完成：非seg {chunkNonSegs.Count} 个，seg {chunkSegs.Count} 个");

            // 对每个 seg 进行分句，然后查找对应的句子名字，并打印每个 seg 的具体句子
            var chunkSegsSentenceNames = new List<List<string>>();
            int segIndex = 0;
            
            foreach (var segContent in chunkSegs)
            {
                segIndex++;
                string escapedSegContent = TextUtils.EscapeNewlinesForDb(segContent);
                bool isSegEmpty = string.IsNullOrEmpty(segContent);
                bool isSegWhitespaceOnly = string.IsNullOrWhiteSpace(segContent);
                ExtractDetailLog(detailLog, $"   处理seg {segIndex}，原始内容长度: {segContent.Length} 字符");
                ExtractDetailLog(detailLog, $"   seg {segIndex} 是否为空: {isSegEmpty}, 是否只包含空白字符: {isSegWhitespaceOnly}");
                ExtractDetailLog(detailLog, $"   seg {segIndex} 内容（已转义）: {escapedSegContent}");
                
                // 对 seg 进行分句
                var segSentences = SplitIntoSentences(segContent);
                ExtractDetailLog(detailLog, $"   seg {segIndex} 分句后共 {segSentences.Count} 个句子");
                
                // 如果 seg 为空或只包含空白字符，且分句后没有句子，尝试直接查找 seg 内容本身
                if (segSentences.Count == 0 && (isSegEmpty || isSegWhitespaceOnly))
                {
                    ExtractDetailLog(detailLog, $"   ⚠️ seg {segIndex} 为空或只包含空白字符，且分句后没有句子，尝试直接查找 seg 内容本身");
                    // 将 seg 内容本身作为一个"句子"来处理
                    segSentences = new List<string> { segContent };
                }

                // 打印每个句子的详细信息（仅调试）
                if (detailLog)
                {
                    for (int i = 0; i < segSentences.Count; i++)
                    {
                        string sentenceContent = segSentences[i];
                        string escapedSentence = TextUtils.EscapeNewlinesForDb(sentenceContent);
                        bool isWhitespaceOnly = string.IsNullOrWhiteSpace(sentenceContent);
                        ExtractDetailLog(true, $"     句子 {i + 1}: 长度={sentenceContent.Length}, 空白字符={isWhitespaceOnly}, 内容（已转义）={escapedSentence}");
                    }
                }

                // 查找每个句子对应的名字
                var segSentenceNames = new List<string>();
                int foundCount = 0;
                int notFoundCount = 0;
                foreach (var sentenceContent in segSentences)
                {
                    // 跳过空字符串，不查找名字（因为空字符串不存储到映射表中）
                    if (string.IsNullOrEmpty(sentenceContent))
                    {
                        ExtractDetailLog(detailLog, $"     跳过空字符串句子");
                        continue;
                    }
                    
                    // 通过句子内容的hash查找对应的名字
                    string contentHash = GenerateContentHash(sentenceContent);
                    ExtractDetailLog(detailLog, $"     查找句子，hash={contentHash.Substring(0, Math.Min(16, contentHash.Length))}...");
                    
                    if (hashToNameMap.TryGetValue(contentHash, out string sentenceName))
                    {
                        segSentenceNames.Add(sentenceName);
                        foundCount++;
                        ExtractDetailLog(detailLog, $"     ✅ 找到句子名称: {sentenceName}");
                    }
                    else
                    {
                        // 如果找不到对应的名字，记录警告（仅调试输出）
                        string escapedSentenceContent = TextUtils.EscapeNewlinesForDb(sentenceContent);
                        bool isWhitespaceOnly = string.IsNullOrWhiteSpace(sentenceContent);
                        ExtractDetailLog(detailLog, $"     ⚠️ 警告：seg {segIndex} 找不到句子内容对应的名字");
                        ExtractDetailLog(detailLog, $"       hash={contentHash.Substring(0, Math.Min(16, contentHash.Length))}...");
                        ExtractDetailLog(detailLog, $"       完整hash={contentHash}");
                        ExtractDetailLog(detailLog, $"       句子长度={sentenceContent.Length}, 空白字符={isWhitespaceOnly}");
                        ExtractDetailLog(detailLog, $"       句子内容（已转义）: {escapedSentenceContent}");
                        notFoundCount++;
                    }
                }
                ExtractDetailLog(detailLog, $"   seg {segIndex} 找到 {foundCount} 个句子名称，未找到 {notFoundCount} 个");
                if (segSentenceNames.Count > 0)
                {
                    ExtractDetailLog(detailLog, $"   seg {segIndex} 句子名称列表: {string.Join(",", segSentenceNames)}");
                }
                else
                {
                    ExtractDetailLog(detailLog, $"   seg {segIndex} 句子名称列表: (空)");
                }

                if (verbose && detailLog)
                {
                    Console.WriteLine($"");
                    Console.WriteLine($"==========================================");
                    Console.WriteLine($"Seg {segIndex} 的具体句子:");
                    Console.WriteLine($"==========================================");
                    for (int i = 0; i < segSentences.Count; i++)
                    {
                        string sentenceContent = segSentences[i];
                        string escapedSentence = TextUtils.EscapeNewlinesForDb(sentenceContent);
                        Console.WriteLine($"句子 {i + 1}: {escapedSentence}");
                    }

                    Console.WriteLine($"==========================================");
                }

                chunkSegsSentenceNames.Add(segSentenceNames);
            }

            // 生成 seg_snapshot
            var chunkSegSnapshotParts = new List<string>();
            foreach (var segSentenceNames in chunkSegsSentenceNames)
            {
                // 同一个 seg 的句子名称用逗号连接
                chunkSegSnapshotParts.Add(string.Join(",", segSentenceNames));
            }

            // 不同 seg 之间用分号连接
            string segSnapshotStr = string.Join(";", chunkSegSnapshotParts);
            
            ExtractDetailLog(detailLog, $"✅ 快照生成完成");
            ExtractDetailLog(detailLog, $"   原快照: {snapshotStr.Substring(0, Math.Min(100, snapshotStr.Length))}...");
            ExtractDetailLog(detailLog, $"   seg快照: {segSnapshotStr.Substring(0, Math.Min(100, segSnapshotStr.Length))}...");

            return (snapshotStr, segSnapshotStr);
        }

        /// <summary>
        /// 去除 display_content 中的句子标记（如 [S_00000-start] 和 [S_00000-end]）
        /// </summary>
        /// <param name="displayContent">包含句子标记的 display_content</param>
        /// <returns>去除标记后的内容</returns>
        private static string RemoveSentenceMarkers(string displayContent)
        {
            if (string.IsNullOrEmpty(displayContent))
            {
                return displayContent;
            }

            // 使用正则表达式匹配 [S_00000-start] 和 [S_00000-end] 格式的标记（S_前缀 + 5位编码：数字、小写字母、大写字母）
            // 模式：\[S_[0-9a-zA-Z]{5}-start\]|\[S_[0-9a-zA-Z]{5}-end\] 匹配 [S_00000-start] 和 [S_00000-end] 这样的标记
            var markerPattern = new System.Text.RegularExpressions.Regex(@"\[S_[0-9a-zA-Z]{5}-start\]|\[S_[0-9a-zA-Z]{5}-end\]");
            return markerPattern.Replace(displayContent, "");
        }

        /// <summary>
        /// 验证 display_content 去除标记后是否与 readText 一致
        /// </summary>
        /// <param name="displayContent">包含句子标记的 display_content</param>
        /// <param name="readText">处理后的原始文本</param>
        /// <param name="chunkIndex">chunk 索引（用于打印）</param>
        private static void VerifyDisplayContent(string displayContent, string readText, int chunkIndex)
        {
            // 去除 display_content 中的 [S_00000-start] 和 [S_00000-end] 类似标记
            string displayContentWithoutMarkers = RemoveSentenceMarkers(displayContent);
            bool isMatch = displayContentWithoutMarkers == readText;
            
            // 打印对比结果
            Console.WriteLine($"");
            Console.WriteLine($"对比结果 (Chunk {chunkIndex}):");
            Console.WriteLine($"  去除标记后的 display_content 长度: {displayContentWithoutMarkers.Length}");
            Console.WriteLine($"  processed_text 长度: {readText.Length}");
            Console.WriteLine($"  是否一致: {(isMatch ? "✅ 一致" : "❌ 不一致")}");
            
            if (!isMatch)
            {
                // 找出不一致的位置
                int minLength = Math.Min(displayContentWithoutMarkers.Length, readText.Length);
                int diffIndex = -1;
                for (int i = 0; i < minLength; i++)
                {
                    if (displayContentWithoutMarkers[i] != readText[i])
                    {
                        diffIndex = i;
                        break;
                    }
                }
                
                if (diffIndex >= 0)
                {
                    int startPos = Math.Max(0, diffIndex - 50);
                    int endPos = Math.Min(minLength, diffIndex + 50);
                    string displayPreview = TextUtils.EscapeNewlinesForDb(displayContentWithoutMarkers.Substring(startPos, endPos - startPos));
                    string processedPreview = TextUtils.EscapeNewlinesForDb(readText.Substring(startPos, endPos - startPos));
                    Console.WriteLine($"  第一个不一致位置: {diffIndex}");
                    Console.WriteLine($"  display_content 片段: ...{displayPreview}...");
                    Console.WriteLine($"  processed_text 片段: ...{processedPreview}...");
                }
                else if (displayContentWithoutMarkers.Length != readText.Length)
                {
                    Console.WriteLine($"  长度不一致，差异: {Math.Abs(displayContentWithoutMarkers.Length - readText.Length)} 字符");
                }
            }
            
            System.Diagnostics.Debug.WriteLine($"对比结果 (Chunk {chunkIndex}): {(isMatch ? "一致" : "不一致")}");
        }

        /// <summary>
        /// 构建 display_content
        /// </summary>
        /// <param name="nonSegs">非seg内容列表</param>
        /// <param name="segsSentenceNames">每个 seg 的句子名称列表</param>
        /// <param name="nameToContentMap">句子名称到内容的映射字典</param>
        /// <returns>display_content 字符串</returns>
        private static string BuildDisplayContent(
            List<string> nonSegs,
            List<List<string>> segsSentenceNames,
            Dictionary<string, string> nameToContentMap)
        {
            var displayParts = new List<string>();

            // 非seg和seg交替出现：非seg1, seg1, 非seg2, seg2, ...
            for (int segIdx = 0; segIdx < segsSentenceNames.Count; segIdx++)
            {
                // 添加非seg内容（在seg之前）
                if (segIdx < nonSegs.Count)
                {
                    displayParts.Add(nonSegs[segIdx]);
                }

                // 构建该 seg 的显示内容（通过句子名称从映射表中查找内容）
                var segDisplayParts = new List<string>();
                foreach (var sentenceName in segsSentenceNames[segIdx])
                {
                    if (nameToContentMap.TryGetValue(sentenceName, out string sentenceContent))
                    {
                        // 保留原始内容（包括换行符），以便与 readText 保持一致
                        // 注意：映射表中的内容已经保留了原始格式（包括末尾的 \r）
                        // 使用 [S_00000-start] 和 [S_00000-end] 格式包裹句子，使大模型更容易识别句子边界
                        segDisplayParts.Add($"[{sentenceName}-start]{sentenceContent}[{sentenceName}-end]");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"⚠️ 警告：找不到句子名称 {sentenceName} 对应的内容");
                    }
                }

                displayParts.Add(string.Join("", segDisplayParts));
            }

            // 添加最后一个非seg（如果有）
            if (nonSegs.Count > segsSentenceNames.Count)
            {
                displayParts.Add(nonSegs[nonSegs.Count - 1]);
            }

            return string.Join("", displayParts);
        }

        /// <summary>
        /// 解析 API 响应并更新 DocumentState（须在 Word STA 线程调用，不可在 async 续延线程中访问 COM）。
        /// </summary>
        private static ProcessingResult ApplyProcessDocumentApiResponse(
            Word.Document document,
            string docUuid,
            string snapshotSource,
            BackendApiClient.JsonResult apiResult)
        {
            try
            {
                if (!apiResult.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"⚠️ 接口请求失败: {apiResult.StatusCode}");
                    if (!string.IsNullOrEmpty(apiResult.Body))
                    {
                        int previewLen = Math.Min(500, apiResult.Body.Length);
                        System.Diagnostics.Debug.WriteLine(
                            $"⚠️ 错误内容（前{previewLen}字符）: {apiResult.Body.Substring(0, previewLen)}");
                    }

                    return null;
                }

                string responseContent = apiResult.Body ?? "";
                System.Diagnostics.Debug.WriteLine($"[API] 响应状态码: {apiResult.StatusCode}");
                System.Diagnostics.Debug.WriteLine($"[API] 响应体长度: {responseContent.Length} 字符");

                var responseObj = JObject.Parse(responseContent);
                if (responseObj["success"]?.Value<bool>() == true && responseObj["data"] != null)
                {
                                var data = responseObj["data"];
                                
                                // 更新DocumentState
                                string returnedDocUuid = data["DocUuid"]?.Value<string>() ?? docUuid;
                                DocumentState.SetCurrDocUuidForDocument(document, returnedDocUuid);
                                
                                // 设置SentenceMappings
                                // 四元组格式：[句子名称, 哈希码, 内容, 标记]
                                // 标记：1表示来自cell类型的seg（cell类型且长度<150），0表示其他类型
                                if (data["SentenceMappings"] != null)
                                {
                                    var sentenceMappings = new List<List<string>>();
                                    foreach (var mapping in data["SentenceMappings"])
                                    {
                                        if (mapping is JArray mappingArray)
                                        {
                                            if (mappingArray.Count >= 4)
                                            {
                                                sentenceMappings.Add(new List<string>
                                                {
                                                    mappingArray[0].Value<string>(),
                                                    mappingArray[1].Value<string>(),
                                                    mappingArray[2].Value<string>(),
                                                    mappingArray[3].Value<string>() ?? "0"
                                                });
                                            }
                                            else
                                            {
                                                // 如果格式不正确，记录警告但跳过
                                                System.Diagnostics.Debug.WriteLine($"[API] ⚠️ 警告：句子映射格式不正确，期望四元组，实际收到 {mappingArray.Count} 个元素，已跳过");
                                            }
                                        }
                                    }
                                    DocumentState.SetSentenceMappings(sentenceMappings);
                                    System.Diagnostics.Debug.WriteLine($"[API] 已设置 {sentenceMappings.Count} 条句子映射（四元组格式：名称、哈希、内容、标记）");
                                }

                                if (data["ParagraphMappings"] != null)
                                {
                                    var paragraphMappings = new List<List<string>>();
                                    foreach (var mapping in data["ParagraphMappings"])
                                    {
                                        if (mapping is JArray mappingArray && mappingArray.Count >= 4)
                                        {
                                            paragraphMappings.Add(new List<string>
                                            {
                                                mappingArray[0].Value<string>(),
                                                mappingArray[1].Value<string>(),
                                                mappingArray[2].Value<string>(),
                                                mappingArray[3].Value<string>() ?? "0"
                                            });
                                        }
                                    }

                                    DocumentState.SetParagraphMappings(paragraphMappings);
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[API] 已设置 {paragraphMappings.Count} 条段落映射");
                                }

                                var s2p = data["SentenceToParagraph"]?.ToObject<Dictionary<string, string>>();
                                var p2s = data["ParagraphToSentences"]?.ToObject<Dictionary<string, List<string>>>();
                                if (s2p != null || p2s != null)
                                {
                                    DocumentState.ApplySentenceParagraphIndexFromApi(s2p, p2s);
                                }

                                var paraDisplays = new List<string>();
                                if (data["ChunkInfo"] != null)
                                {
                                    foreach (var chunkData in data["ChunkInfo"])
                                    {
                                        string pd = chunkData["paragraph_display_content"]?.Value<string>() ?? "";
                                        paraDisplays.Add(pd);
                                    }
                                }

                                if (paraDisplays.Count > 0)
                                {
                                    DocumentState.SetChunkParagraphDisplayContents(paraDisplays);
                                }
                                
                                // 设置Snapshot和SegSnapshot
                                string snapshot = data["Snapshot"]?.Value<string>() ?? "";
                                string segSnapshot = data["SegSnapshot"]?.Value<string>() ?? "";
                                string historySource = string.IsNullOrEmpty(snapshotSource) ? "process_document_api" : snapshotSource;
                                DocumentState.SetSnapshots(snapshot, segSnapshot, historySource);
                                System.Diagnostics.Debug.WriteLine($"[API] 已设置快照: Snapshot={snapshot}, SegSnapshot={segSnapshot}");

                                // 设置SnapshotIdx（如果API返回了该字段）
                                // 如果没有收到API返回或字段不存在，保持默认值-1
                                if (data["SnapshotIdx"] != null)
                                {
                                    int snapshotIdx = data["SnapshotIdx"]?.Value<int>() ?? -1;
                                    DocumentState.SetSnapshotIdx(snapshotIdx);
                                    System.Diagnostics.Debug.WriteLine($"[API] 已设置快照索引: {snapshotIdx}");
                                }
                                else
                                {
                                    // 明确设置为-1，表示未收到API返回
                                    DocumentState.SetSnapshotIdx(-1);
                                    System.Diagnostics.Debug.WriteLine($"[API] 未收到快照索引，设置为默认值-1");
                                }

                                // 从API返回的数据中提取并设置TableIndexToIdMap
                                if (data["TableIndexToIdMap"] != null)
                                {
                                    var tableIndexToIdMapFromApi = data["TableIndexToIdMap"]?.ToObject<Dictionary<int, string>>();
                                    if (tableIndexToIdMapFromApi != null)
                                    {
                                        DocumentState.SetTableIndexToIdMap(tableIndexToIdMapFromApi);
                                        System.Diagnostics.Debug.WriteLine($"[API] 已设置表格序号到编号映射，共 {tableIndexToIdMapFromApi.Count} 个表格");
                                    }
                                }

                                if (data["ChartIndexToIdMap"] != null)
                                {
                                    var chartIndexToIdMapFromApi = data["ChartIndexToIdMap"]?.ToObject<Dictionary<int, string>>();
                                    if (chartIndexToIdMapFromApi != null)
                                    {
                                        DocumentState.SetChartIdOrder(
                                            chartIndexToIdMapFromApi.Keys.OrderBy(k => k)
                                                .Select(k => chartIndexToIdMapFromApi[k]).ToList());
                                        System.Diagnostics.Debug.WriteLine(
                                            $"[API] 已设置图表序号到编号映射，共 {chartIndexToIdMapFromApi.Count} 个图表");
                                    }
                                }

                                if (data["ImageIndexToIdMap"] != null)
                                {
                                    var imageIndexToIdMapFromApi = data["ImageIndexToIdMap"]?.ToObject<Dictionary<int, string>>();
                                    if (imageIndexToIdMapFromApi != null)
                                    {
                                        DocumentState.SetImageIndexToIdMap(imageIndexToIdMapFromApi);
                                        System.Diagnostics.Debug.WriteLine(
                                            $"[API] 已设置图片序号到编号映射，共 {imageIndexToIdMapFromApi.Count} 个图片");
                                    }
                                }

                                // 生成表格序号到编号的映射（备用，如果API没有返回）
                                var tableIndexToIdMap = new Dictionary<int, string>();
                                var tableIdOrder = DocumentState.GetTableIdOrder();
                                for (int i = 0; i < tableIdOrder.Count; i++)
                                {
                                    tableIndexToIdMap[i + 1] = tableIdOrder[i]; // 序号从1开始
                                }
                                System.Diagnostics.Debug.WriteLine($"✅ API处理：生成表格序号到编号映射，共 {tableIndexToIdMap.Count} 个表格");

                                // 构建ProcessingResult
                                var processingResult = new ProcessingResult
                                {
                                    ReadText = data["ReadText"]?.Value<string>() ?? "",
                                    Snapshot = snapshot,
                                    SegSnapshot = segSnapshot,
                                    NameToContentMap = data["NameToContentMap"]?.ToObject<Dictionary<string, string>>() ?? new Dictionary<string, string>(),
                                    ChunkInfo = new List<ChunkInfo>(),
                                    TableIndexToIdMap = tableIndexToIdMap,
                                    ImageIndexToIdMap = DocumentState.GetImageIdOrder()
                                        .Select((id, idx) => new { idx, id })
                                        .ToDictionary(x => x.idx, x => x.id),
                                    ProcessingMethod = "API"
                                };
                                
                                // 转换ChunkInfo
                                if (data["ChunkInfo"] != null)
                                {
                                    foreach (var chunkData in data["ChunkInfo"])
                                    {
                                        processingResult.ChunkInfo.Add(new ChunkInfo
                                        {
                                            ChunkIdx = chunkData["chunk_idx"]?.Value<int>() ?? 0,
                                            Content = chunkData["content"]?.Value<string>() ?? "",
                                            DisplayContent = chunkData["display_content"]?.Value<string>() ?? "",
                                            ParagraphDisplayContent = chunkData["paragraph_display_content"]?.Value<string>() ?? "",
                                            Summary = chunkData["summary"]?.Value<string>() ?? ""
                                        });
                                    }
                                }
                                
                                System.Diagnostics.Debug.WriteLine("✅ 接口处理文档成功，DocumentState已更新");
                                return processingResult;
                }
                System.Diagnostics.Debug.WriteLine($"⚠️ 接口返回失败: {responseContent}");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"⚠️ 解析接口响应时发生异常: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"⚠️ 堆栈跟踪: {ex.StackTrace}");
                return null;
            }
        }

        /// <summary>
        /// 轻量 API 路径：另存副本供 multipart 上传（不切换活动文档，避免 Word 占用文件）。
        /// 复用 <see cref="DocumentCheckpointSaveHelper.SaveDocumentCopy"/>。
        /// </summary>
        private static bool TrySaveDocumentCopyForApi(
            Word.Document document,
            bool verbose,
            out string docCopyPath,
            out string tempFilePath)
        {
            docCopyPath = null;
            tempFilePath = null;

            try
            {
                var userService = UserService.Instance;
                string userName = userService.IsLoggedIn ? userService.UserName : "anonymous";
                string baseDir = McpToolsHelpers.GetUserWorkDirectory("curr_docs", userName);
                if (string.IsNullOrEmpty(baseDir))
                {
                    return false;
                }

                string docName = string.IsNullOrEmpty(document.Name) ? "文档" : Path.GetFileNameWithoutExtension(document.Name);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string tempFileName = $"{timestamp}_{docName}.docx";
                tempFilePath = Path.Combine(baseDir, tempFileName);

                DocumentCheckpointSaveHelper.SaveDocumentCopy(document, tempFilePath);
                docCopyPath = tempFilePath;

                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine($"[API副本] ✅ SaveDocumentCopy: {docCopyPath}");
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[API副本] ⚠️ SaveDocumentCopy 失败: {ex.Message}");
                docCopyPath = null;
                tempFilePath = null;
                return false;
            }
        }

        /// <summary>
        /// SaveAs 文档到用户 curr_docs 目录（ExtractImages 路径：Word 占用 SaveAs 路径，须另建 _copy 供解压/读取）。
        /// </summary>
        private static bool TrySaveDocumentCopies(
            Word.Document document,
            bool createSeparateCopy,
            bool verbose,
            out string docCopyPath,
            out string tempFilePath)
        {
            docCopyPath = null;
            tempFilePath = null;

            try
            {
                var userService = UserService.Instance;
                string userName = userService.IsLoggedIn ? userService.UserName : "anonymous";
                string baseDir = McpToolsHelpers.GetUserWorkDirectory("curr_docs", userName);
                if (string.IsNullOrEmpty(baseDir))
                {
                    return false;
                }

                string docName = string.IsNullOrEmpty(document.Name) ? "文档" : Path.GetFileNameWithoutExtension(document.Name);
                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string tempFileName = $"{timestamp}_{docName}.docx";
                tempFilePath = Path.Combine(baseDir, tempFileName);

                document.SaveAs2(tempFilePath, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
                if (verbose)
                {
                    System.Diagnostics.Debug.WriteLine($"[第一步半] ✅ 文档副本已保存: {tempFilePath}");
                }

                if (createSeparateCopy)
                {
                    System.Threading.Thread.Sleep(500);

                    string copyFileName = Path.GetFileNameWithoutExtension(tempFileName) + "_copy.docx";
                    string copyFilePath = Path.Combine(baseDir, copyFileName);
                    File.Copy(tempFilePath, copyFilePath, overwrite: true);
                    docCopyPath = copyFilePath;
                    if (verbose)
                    {
                        System.Diagnostics.Debug.WriteLine($"[第一步半] ✅ 文档副本的副本已创建: {docCopyPath}");
                    }
                }
                else
                {
                    docCopyPath = tempFilePath;
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[文档副本] ⚠️ SaveAs 失败: {ex.Message}");
                docCopyPath = null;
                tempFilePath = null;
                return false;
            }
        }

        /// <summary>
        /// 清理临时文件
        /// </summary>
        /// <param name="docCopyPath">文档副本的副本路径</param>
        /// <param name="tempFilePath">临时文件路径</param>
        private static void CleanupTempFiles(string docCopyPath, string tempFilePath)
        {
            // 删除临时文件副本的副本
            if (!string.IsNullOrEmpty(docCopyPath) && File.Exists(docCopyPath))
            {
                try
                {
                    // 等待一下，确保文件没有被占用
                    System.Threading.Thread.Sleep(500);
                    File.Delete(docCopyPath);
                    System.Diagnostics.Debug.WriteLine($"[清理] ✅ 临时文件副本的副本已删除: {docCopyPath}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[清理] ⚠️ 删除临时文件副本的副本失败: {ex.Message}");
                }
            }
            
            // 删除原始临时文件副本
            if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
            {
                try
                {
                    // 等待一下，确保文件没有被占用
                    System.Threading.Thread.Sleep(500);
                    File.Delete(tempFilePath);
                    System.Diagnostics.Debug.WriteLine($"[清理] ✅ 临时文件副本已删除: {tempFilePath}");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[清理] ⚠️ 删除临时文件副本失败: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Chunk信息
        /// </summary>
        public class ChunkInfo
        {
            public int ChunkIdx { get; set; }
            public string Content { get; set; }
            public string DisplayContent { get; set; }
            public string ParagraphDisplayContent { get; set; }
            public string Summary { get; set; }
        }

        /// <summary>
        /// 处理结果
        /// </summary>
        public class ProcessingResult
        {
            public string ReadText { get; set; }
            public List<ChunkInfo> ChunkInfo { get; set; }
            public Dictionary<string, string> NameToContentMap { get; set; }
            public string Snapshot { get; set; }
            public string SegSnapshot { get; set; }
            public Dictionary<int, string> TableIndexToIdMap { get; set; }
            public Dictionary<int, string> ImageIndexToIdMap { get; set; }
            /// <summary>
            /// 处理方式：API 表示使用API处理，Local 表示使用本地处理
            /// </summary>
            public string ProcessingMethod { get; set; }
        }
    }
}




