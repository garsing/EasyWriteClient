using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 文档状态管理类，用于存储和管理文档编辑操作的状态信息
    /// 历史记录：大模型已经执行格式修改（添加删除线、黄色背景等）但用户未确认的记录
    /// 用户确认后，这些记录应该被删除
    /// </summary>
    public static class DocumentState
    {
        // 历史的新增和删除的range（已执行格式修改但未确认的操作）
        private static readonly List<(int Start, int End)> _historicalInsertRanges = new List<(int Start, int End)>();
        private static readonly List<(int Start, int End)> _historicalDeleteRanges = new List<(int Start, int End)>();

        // 添加按钮后的最终删除和插入范围（用于内容读取过滤）
        private static readonly List<(int Start, int End)> _finalDeleteRanges = new List<(int Start, int End)>();
        private static readonly List<(int Start, int End)> _finalInsertRanges = new List<(int Start, int End)>();

        // 段落跟踪相关字段
        // 映射表格式：[['AAA','哈希码','内容'],['AAB','哈希码','内容']]
        // 每个元素是 [段落名称, 内容哈希码, 段落内容]
        private static readonly List<List<string>> _paragraphNameMapping = new List<List<string>>();
        private static readonly List<List<string>> _snapshots = new List<List<string>>();
        private static int _nextNameIndex = 0;
        private static int _oldSnapshotIndex = 0; // old_snapshot指针，指向用于比对的快照，默认指向最初快照（索引0）
        private static List<(string header, List<string> lines)> _displayHunk = new List<(string header, List<string> lines)>(); // display_hunk：记录当前Word中显示的全量hunk（包含所有段落，header为空字符串）

        // 当前文档UUID
        private static string _currDocUuid = null;

        /// <summary>
        /// 获取历史新增ranges的只读副本
        /// </summary>
        public static IReadOnlyList<(int Start, int End)> HistoricalInsertRanges => _historicalInsertRanges.AsReadOnly();

        /// <summary>
        /// 获取历史删除ranges的只读副本
        /// </summary>
        public static IReadOnlyList<(int Start, int End)> HistoricalDeleteRanges => _historicalDeleteRanges.AsReadOnly();

        /// <summary>
        /// 获取最终删除ranges的只读副本（添加按钮后的实际范围）
        /// </summary>
        public static IReadOnlyList<(int Start, int End)> FinalDeleteRanges => _finalDeleteRanges.AsReadOnly();

        /// <summary>
        /// 获取最终插入ranges的只读副本（添加按钮后的实际范围）
        /// </summary>
        public static IReadOnlyList<(int Start, int End)> FinalInsertRanges => _finalInsertRanges.AsReadOnly();

        /// <summary>
        /// 获取段落命名映射表
        /// 格式：[['AAA','哈希码','内容'],['AAB','哈希码','内容']]
        /// </summary>
        public static IReadOnlyList<IReadOnlyList<string>> ParagraphNameMapping => 
            _paragraphNameMapping.Select(m => (IReadOnlyList<string>)m.AsReadOnly()).ToList().AsReadOnly();

        /// <summary>
        /// 获取所有快照历史
        /// </summary>
        public static IReadOnlyList<IReadOnlyList<string>> Snapshots => _snapshots.Select(s => s.AsReadOnly()).ToList().AsReadOnly();

        /// <summary>
        /// 获取初始快照
        /// </summary>
        public static IReadOnlyList<string> InitialSnapshot => _snapshots.Count > 0 ? _snapshots[0].AsReadOnly() : new List<string>().AsReadOnly();

        /// <summary>
        /// 获取当前快照（最新的快照）
        /// </summary>
        public static IReadOnlyList<string> CurrentSnapshot => _snapshots.Count > 0 ? _snapshots[_snapshots.Count - 1].AsReadOnly() : new List<string>().AsReadOnly();

        /// <summary>
        /// 获取old_snapshot指针指向的快照（用于和最新快照做比对）
        /// </summary>
        public static IReadOnlyList<string> OldSnapshot => _snapshots.Count > 0 && _oldSnapshotIndex >= 0 && _oldSnapshotIndex < _snapshots.Count 
            ? _snapshots[_oldSnapshotIndex].AsReadOnly() 
            : new List<string>().AsReadOnly();

        /// <summary>
        /// 获取old_snapshot指针的索引
        /// </summary>
        public static int OldSnapshotIndex => _oldSnapshotIndex;

        /// <summary>
        /// 获取display_hunk（当前Word中显示的全量hunk，包含所有段落）
        /// </summary>
        public static IReadOnlyList<(string header, IReadOnlyList<string> lines)> DisplayHunk => 
            _displayHunk.Select(h => (h.header, (IReadOnlyList<string>)h.lines.AsReadOnly())).ToList().AsReadOnly();

        /// <summary>
        /// 获取当前文档UUID
        /// </summary>
        public static string CurrDocUuid => _currDocUuid;

        /// <summary>
        /// 设置当前文档UUID
        /// </summary>
        /// <param name="docUuid">文档UUID</param>
        public static void SetCurrDocUuid(string docUuid)
        {
            _currDocUuid = docUuid;
            System.Diagnostics.Debug.WriteLine($"[DocumentState] 设置当前文档UUID: {docUuid}");
        }

        /// <summary>
        /// 清除当前文档UUID
        /// </summary>
        public static void ClearCurrDocUuid()
        {
            _currDocUuid = null;
            System.Diagnostics.Debug.WriteLine("[DocumentState] 清除当前文档UUID");
        }

        /// <summary>
        /// 获取原始的完整Word显示内容（从当前的display_hunk直接获取，全量形式）
        /// 这是起点（originalFullDisplayContent），用于对照
        /// 全量形式：直接从hunk中去掉-和+符号即可
        /// 如果display_hunk为空，则从old_snapshot初始化display_hunk
        /// </summary>
        public static List<string> GetOriginalFullDisplayContent()
        {
            // 如果display_hunk为空，从old_snapshot初始化display_hunk
            if (_displayHunk == null || _displayHunk.Count == 0)
            {
                var oldSnapshot = OldSnapshot.ToList();
                if (oldSnapshot.Count > 0)
                {
                    // 将old_snapshot转化为全量hunk格式（所有段落都是未改动的，不需要-或+前缀）
                    var fullHunkLines = new List<string>();
                    foreach (var paraName in oldSnapshot)
                    {
                        fullHunkLines.Add(paraName);
                    }
                    
                    // 初始化display_hunk
                    _displayHunk = new List<(string header, List<string> lines)>
                    {
                        ("", fullHunkLines)
                    };
                    
                    System.Diagnostics.Debug.WriteLine($"[段落跟踪] display_hunk为空，已从old_snapshot初始化，包含 {fullHunkLines.Count} 个段落");
                }
            }
            
            return GetFullDisplayContentFromHunk(_displayHunk);
        }

        /// <summary>
        /// 获取目标的完整Word显示内容（从新的全量hunk直接获取）
        /// 这是终点（targetFullDisplayContent），是Word中要展示的最终状态
        /// 全量形式：直接从hunk中去掉-和+符号即可
        /// </summary>
        /// <param name="hunk">新的全量hunk列表（只包含一个全量hunk）</param>
        /// <returns>目标显示内容的段落名列表</returns>
        public static List<string> GetTargetFullDisplayContent(List<(string header, List<string> lines)> hunk)
        {
            return GetFullDisplayContentFromHunk(hunk);
        }

        /// <summary>
        /// 通用的方法：根据全量hunk生成完整的显示内容
        /// 全量形式：直接从hunk中去掉-和+符号即可
        /// </summary>
        /// <param name="hunk">全量hunk列表（只包含一个全量hunk，header为空，lines包含所有段落）</param>
        /// <returns>合成后的完整显示内容</returns>
        private static List<string> GetFullDisplayContentFromHunk(List<(string header, List<string> lines)> hunk)
        {
            var result = new List<string>();

            // 全量形式：直接从hunk中去掉-和+符号
            foreach (var (header, lines) in hunk)
            {
                foreach (var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                    {
                        if (line.StartsWith("-") || line.StartsWith("+"))
                        {
                            // 去掉-或+符号，只保留段落名称
                            result.Add(line.Substring(1));
                        }
                        else
                        {
                            // 未改动的段落，直接添加
                            result.Add(line);
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 添加历史新增range（当执行格式修改时调用）
        /// </summary>
        public static void AddHistoricalInsertRange(int start, int end)
        {
            _historicalInsertRanges.Add((start, end));
        }

        /// <summary>
        /// 添加历史删除range（当执行格式修改时调用）
        /// </summary>
        public static void AddHistoricalDeleteRange(int start, int end)
        {
            _historicalDeleteRanges.Add((start, end));
        }

        /// <summary>
        /// 移除指定的历史新增range（当操作被确认时调用）
        /// </summary>
        public static void RemoveHistoricalInsertRange(int start, int end)
        {
            _historicalInsertRanges.Remove((start, end));
        }

        /// <summary>
        /// 移除指定的历史删除range（当操作被确认时调用）
        /// </summary>
        public static void RemoveHistoricalDeleteRange(int start, int end)
        {
            _historicalDeleteRanges.Remove((start, end));
        }

        /// <summary>
        /// 清除所有历史ranges（重置时调用）
        /// </summary>
        public static void ClearAll()
        {
            _historicalInsertRanges.Clear();
            _historicalDeleteRanges.Clear();
            _finalDeleteRanges.Clear();
            _finalInsertRanges.Clear();
            _paragraphNameMapping.Clear();
            _snapshots.Clear();
            _nextNameIndex = 0;
        }

        /// <summary>
        /// 添加最终删除range（添加按钮后的实际范围）
        /// </summary>
        public static void AddFinalDeleteRange(int start, int end)
        {
            _finalDeleteRanges.Add((start, end));
        }

        /// <summary>
        /// 添加最终插入range（添加按钮后的实际范围）
        /// </summary>
        public static void AddFinalInsertRange(int start, int end)
        {
            _finalInsertRanges.Add((start, end));
        }

        /// <summary>
        /// 清除所有最终ranges
        /// </summary>
        public static void ClearFinalRanges()
        {
            _finalDeleteRanges.Clear();
            _finalInsertRanges.Clear();
        }

        /// <summary>
        /// 获取状态统计信息
        /// </summary>
        public static (int historicalInserts, int historicalDeletes, int finalInserts, int finalDeletes) GetStats()
        {
            return (
                _historicalInsertRanges.Count,
                _historicalDeleteRanges.Count,
                _finalInsertRanges.Count,
                _finalDeleteRanges.Count
            );
        }

        /// <summary>
        /// 获取整合后的历史修改记录
        /// 将所有历史修改整理成删除和新增的列表形式，区间从小到大排列
        /// </summary>
        /// <returns>包含删除和新增区间的元组</returns>
        public static (List<(int Start, int End)> consolidatedDeletes, List<(int Start, int End)> consolidatedInserts) GetConsolidatedHistoricalRanges()
        {
            // 创建副本并排序
            var deletes = new List<(int Start, int End)>(_historicalDeleteRanges);
            var inserts = new List<(int Start, int End)>(_historicalInsertRanges);

            // 按起始位置从小到大排序
            deletes.Sort((a, b) => a.Start.CompareTo(b.Start));
            inserts.Sort((a, b) => a.Start.CompareTo(b.Start));

            return (deletes, inserts);
        }

        #region 段落跟踪功能

        /// <summary>
        /// 生成内容的哈希码（用于验证内容唯一性）
        /// </summary>
        private static string GenerateContentHash(string content)
        {
            if (string.IsNullOrEmpty(content))
                return "";
            
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(content));
                StringBuilder sb = new StringBuilder();
                foreach (byte b in hashBytes)
                {
                    sb.Append(b.ToString("x2"));
                }
                return sb.ToString();
            }
        }

        /// <summary>
        /// 根据内容查找或创建段落名称
        /// 如果内容已存在（通过哈希码验证），返回已存在的名称；否则创建新名称
        /// </summary>
        private static string FindOrCreateParagraphName(string content)
        {
            string contentHash = GenerateContentHash(content);
            
            // 查找是否已存在相同内容的段落（通过哈希码验证）
            foreach (var mapping in _paragraphNameMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    // 内容已存在，返回已存在的名称
                    return mapping[0];
                }
            }
            
            // 内容不存在，创建新名称
            string newName = GenerateNextName();
            _paragraphNameMapping.Add(new List<string> { newName, contentHash, content });
            return newName;
        }

        /// <summary>
        /// 生成下一个段落名称 (AAA, AAB, AAC, ...)
        /// </summary>
        private static string GenerateNextName()
        {
            string name = "";
            int index = _nextNameIndex++;

            // 生成三位字母名称 (AAA, AAB, AAC, ..., ZZZ)
            for (int i = 0; i < 3; i++)
            {
                name = (char)('A' + (index % 26)) + name;
                index /= 26;
            }

            return name;
        }

        /// <summary>
        /// 初始化段落跟踪 - 为所有段落分配初始名称并创建初始快照
        /// 允许快照中同一个名字的段落重复排列（例如 AAA,AAB,AAC,AAB,AAD）
        /// </summary>
        /// <param name="paragraphs">段落内容列表</param>
        public static void InitializeParagraphTracking(List<string> paragraphs)
        {
            ClearParagraphTracking();

            System.Diagnostics.Debug.WriteLine("[段落跟踪第一步] 初始化段落跟踪...");

            var initialSnapshot = new List<string>();
            for (int i = 0; i < paragraphs.Count; i++)
            {
                string content = paragraphs[i] ?? "";
                // 根据内容查找或创建段落名称（保证内容是唯一的，通过哈希码验证）
                string name = FindOrCreateParagraphName(content);
                initialSnapshot.Add(name);

                System.Diagnostics.Debug.WriteLine($"[段落跟踪] 段落{i + 1} -> {name}: {content.Substring(0, Math.Min(50, content.Length))}...");
            }

            _snapshots.Add(initialSnapshot);

            // 初始化display_hunk：与初始快照一致，所有段落都是未改动的（没有-或+前缀）
            // display_hunk格式：全量hunk，header为空字符串，lines包含所有段落名称（未改动行）
            _displayHunk = new List<(string header, List<string> lines)>
            {
                ("", new List<string>(initialSnapshot))
            };

            System.Diagnostics.Debug.WriteLine($"[段落跟踪] 初始快照已创建: {string.Join(",", initialSnapshot)}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪] display_hunk已初始化，包含 {initialSnapshot.Count} 个段落（与初始快照一致，都是未改动行）");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪] 命名映射表: {GetNameMappingString()}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪] 快照历史: {GetSnapshotsString()}");
        }

        /// <summary>
        /// 处理段落替换操作
        /// </summary>
        /// <param name="paragraphIndex">段落索引（从1开始）</param>
        /// <param name="newContent">新内容</param>
        public static void ProcessParagraphReplace(int paragraphIndex, string newContent)
        {
            if (_snapshots.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪] 错误：未初始化段落跟踪");
                return;
            }

            var currentSnapshot = _snapshots[_snapshots.Count - 1];
            if (paragraphIndex < 1 || paragraphIndex > currentSnapshot.Count)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪] 错误：无效的段落索引 {paragraphIndex}");
                return;
            }

            // 创建新的快照
            var newSnapshot = new List<string>(currentSnapshot);

            string oldName = currentSnapshot[paragraphIndex - 1];
            string newName = GenerateNextName();
            string safeNewContent = newContent ?? "";

            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 替换操作：段落{paragraphIndex} ({oldName}) -> 新名称{newName}");
            string oldContent = GetParagraphContent(oldName);
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 原内容: {oldContent.Substring(0, Math.Min(50, oldContent.Length))}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 新内容: {safeNewContent.Substring(0, Math.Min(50, safeNewContent.Length))}");

            // 根据新内容查找或创建段落名称（保证内容是唯一的，通过哈希码验证）
            string actualNewName = FindOrCreateParagraphName(safeNewContent);
            // 如果找到已存在的名称，使用已存在的名称；否则使用新生成的名称
            if (actualNewName != newName)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 新内容已存在，使用已存在的名称: {actualNewName}");
                newName = actualNewName;
            }

            // 更新新快照
            newSnapshot[paragraphIndex - 1] = newName;

            // 添加新快照
            _snapshots.Add(newSnapshot);

            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 新快照已创建: {string.Join(",", newSnapshot)}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 快照历史: {GetSnapshotsString()}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 命名映射表已更新: {GetNameMappingString()}");
        }

        /// <summary>
        /// 处理段落删除操作
        /// </summary>
        /// <param name="paragraphIndex">段落索引（从1开始）</param>
        public static void ProcessParagraphDelete(int paragraphIndex)
        {
            if (_snapshots.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪] 错误：未初始化段落跟踪");
                return;
            }

            var currentSnapshot = _snapshots[_snapshots.Count - 1];
            if (paragraphIndex < 1 || paragraphIndex > currentSnapshot.Count)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪] 错误：无效的段落索引 {paragraphIndex}");
                return;
            }

            // 创建新的快照
            var newSnapshot = new List<string>(currentSnapshot);

            string deletedName = currentSnapshot[paragraphIndex - 1];
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 删除操作：段落{paragraphIndex} ({deletedName})");
            string deletedContent = GetParagraphContent(deletedName);
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 删除内容: {deletedContent.Substring(0, Math.Min(50, deletedContent.Length))}");

            // 从新快照中移除
            newSnapshot.RemoveAt(paragraphIndex - 1);

            // 添加新快照
            _snapshots.Add(newSnapshot);

            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 新快照已创建: {string.Join(",", newSnapshot)}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 快照历史: {GetSnapshotsString()}");
        }

        /// <summary>
        /// 处理段落插入操作
        /// </summary>
        /// <param name="paragraphIndex">插入位置的段落索引（从1开始）</param>
        /// <param name="position">插入位置（"before" 或 "after"）</param>
        /// <param name="content">插入的内容</param>
        public static void ProcessParagraphInsert(int paragraphIndex, string position, string content)
        {
            if (_snapshots.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪] 错误：未初始化段落跟踪");
                return;
            }

            var currentSnapshot = _snapshots[_snapshots.Count - 1];
            if (paragraphIndex < 1 || paragraphIndex > currentSnapshot.Count + 1)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪] 错误：无效的插入位置 {paragraphIndex}");
                return;
            }

            // 创建新的快照
            var newSnapshot = new List<string>(currentSnapshot);

            string safeContent = content ?? "";
            // 根据内容查找或创建段落名称（保证内容是唯一的，通过哈希码验证）
            string newName = FindOrCreateParagraphName(safeContent);
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 插入操作：在段落{paragraphIndex} {position ?? "after"} 插入新段落 {newName}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 插入内容: {safeContent.Substring(0, Math.Min(50, safeContent.Length))}");

            // 计算插入位置
            int insertIndex = (position?.ToLower() ?? "after") == "before" ? paragraphIndex - 1 : paragraphIndex;

            // 插入到新快照
            newSnapshot.Insert(insertIndex, newName);

            // 添加新快照
            _snapshots.Add(newSnapshot);

            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 新快照已创建: {string.Join(",", newSnapshot)}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 快照历史: {GetSnapshotsString()}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪操作] 命名映射表已更新: {GetNameMappingString()}");
        }

        /// <summary>
        /// 计算并输出全量形式的hunk（比较初始快照和最终快照）
        /// 全量形式：一个hunk包含所有段落（未改动的和改动的），不需要hunk头
        /// </summary>
        /// <param name="contextLines">保留参数，已废弃（全量形式不需要上下文行数）</param>
        /// <param name="maxGap">保留参数，已废弃（全量形式只有一个hunk）</param>
        /// <returns>返回全量hunk的列表（只包含一个全量hunk，不包含header，只包含全量段落行）</returns>
        public static List<(string header, List<string> lines)> CalculateAndPrintHunk(int contextLines = 3, int maxGap = 5)
        {
            var hunkList = new List<(string header, List<string> lines)>();
            
            if (_snapshots.Count < 2)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪第二步] ===== Hunk计算 =====");
                System.Diagnostics.Debug.WriteLine("[段落跟踪第二步] 没有足够的快照进行比较");
                return hunkList;
            }

            System.Diagnostics.Debug.WriteLine("[段落跟踪第二步] ===== 计算全量Hunk =====");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪第二步] 总快照数量: {_snapshots.Count}");

            // 使用old_snapshot指针指向的快照作为初始快照，用于和最新快照做比对
            var oldSnapshot = OldSnapshot.ToList();
            var finalSnapshot = _snapshots[_snapshots.Count - 1];

            if (oldSnapshot.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪第二步] old_snapshot为空，无法进行比较");
                return hunkList;
            }

            System.Diagnostics.Debug.WriteLine($"[段落跟踪第二步] old_snapshot索引: {_oldSnapshotIndex}, 快照: {string.Join(",", oldSnapshot)}");
            System.Diagnostics.Debug.WriteLine($"[段落跟踪第二步] 最终快照: {string.Join(",", finalSnapshot)}");

            // 找到最长公共子序列（LCS）作为锚点
            // 基于段落名称进行比较，而不是内容
            // 只有名称相同且在正确位置的段落才能作为锚点
            var lcs = FindLongestCommonSubsequence(oldSnapshot, finalSnapshot);

            System.Diagnostics.Debug.WriteLine($"[段落跟踪第二步] LCS锚点段落（基于名称）: {string.Join(",", lcs)}");

            // 构建全量diff行列表（包含所有段落：未改动的和改动的）
            // 全量形式：一个hunk包含所有段落，不需要hunk头
            var fullDiffLines = new List<string>();

            // 基于LCS锚点，使用标准diff算法生成全量差异
            int oldIndex = 0;
            int newIndex = 0;
            int lcsIdx = 0;

            while (oldIndex < oldSnapshot.Count || newIndex < finalSnapshot.Count)
            {
                if (lcsIdx < lcs.Count)
                {
                    // 找到下一个LCS元素在两个序列中的位置
                    string anchor = lcs[lcsIdx];
                    int oldAnchorPos = oldSnapshot.IndexOf(anchor, oldIndex);
                    int newAnchorPos = finalSnapshot.IndexOf(anchor, newIndex);

                    if (oldAnchorPos >= 0 && newAnchorPos >= 0)
                    {
                        // 添加从当前位置到锚点之前的所有差异
                        // 删除：oldSnapshot中从oldIndex到oldAnchorPos之前的所有元素
                        for (int k = oldIndex; k < oldAnchorPos; k++)
                        {
                            fullDiffLines.Add($"-{oldSnapshot[k]}");
                        }

                        // 新增：finalSnapshot中从newIndex到newAnchorPos之前的所有元素
                        for (int k = newIndex; k < newAnchorPos; k++)
                        {
                            fullDiffLines.Add($"+{finalSnapshot[k]}");
                        }

                        // 添加锚点（未改动的段落，不加前缀）
                        fullDiffLines.Add(anchor);

                        // 更新索引
                        oldIndex = oldAnchorPos + 1;
                        newIndex = newAnchorPos + 1;
                        lcsIdx++;
                        continue;
                    }
                }

                // 如果没有找到锚点，处理剩余元素
                // 删除剩余的oldSnapshot元素
                for (int k = oldIndex; k < oldSnapshot.Count; k++)
                {
                    fullDiffLines.Add($"-{oldSnapshot[k]}");
                }

                // 新增剩余的finalSnapshot元素
                for (int k = newIndex; k < finalSnapshot.Count; k++)
                {
                    fullDiffLines.Add($"+{finalSnapshot[k]}");
                }

                break;
            }

            // 生成全量hunk（不需要header，header设为空字符串）
            if (fullDiffLines.Count > 0)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪第二步] 生成全量Hunk（包含所有段落）:");
                foreach (var line in fullDiffLines)
                {
                    System.Diagnostics.Debug.WriteLine(line);
                }

                // 全量形式：header为空字符串，lines包含所有段落
                hunkList.Add(("", fullDiffLines));
            }

            return hunkList;
        }

        /// <summary>
        /// 根据段落名称获取段落内容
        /// </summary>
        public static string GetParagraphContent(string paragraphName)
        {
            foreach (var mapping in _paragraphNameMapping)
            {
                if (mapping.Count >= 3 && mapping[0] == paragraphName)
                {
                    return mapping[2]; // 返回内容
                }
            }
            return "";
        }

        /// <summary>
        /// 获取所有段落映射
        /// 格式：[['AAA','哈希码','内容'],['AAB','哈希码','内容']]
        /// </summary>
        public static IReadOnlyList<IReadOnlyList<string>> GetAllParagraphMappings()
        {
            return _paragraphNameMapping.Select(m => (IReadOnlyList<string>)m.AsReadOnly()).ToList().AsReadOnly();
        }

        /// <summary>
        /// 根据内容查找段落名称（通过哈希码验证）
        /// 如果内容已存在，返回已存在的名称；否则返回null
        /// </summary>
        public static string FindParagraphNameByContent(string content)
        {
            string contentHash = GenerateContentHash(content);
            
            foreach (var mapping in _paragraphNameMapping)
            {
                if (mapping.Count >= 2 && mapping[1] == contentHash)
                {
                    return mapping[0];
                }
            }
            
            return null;
        }

        /// <summary>
        /// 找到两个列表的最长公共子序列（LCS）
        /// </summary>
        private static List<string> FindLongestCommonSubsequence(List<string> list1, List<string> list2)
        {
            int m = list1.Count;
            int n = list2.Count;

            // 使用动态规划计算LCS长度
            int[,] dp = new int[m + 1, n + 1];

            for (int i = 1; i <= m; i++)
            {
                for (int j = 1; j <= n; j++)
                {
                    if (list1[i - 1] == list2[j - 1])
                    {
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                    }
                }
            }

            // 回溯构建LCS
            var lcs = new List<string>();
            int x = m, y = n;
            while (x > 0 && y > 0)
            {
                if (list1[x - 1] == list2[y - 1])
                {
                    lcs.Insert(0, list1[x - 1]);
                    x--;
                    y--;
                }
                else if (dp[x - 1, y] > dp[x, y - 1])
                {
                    x--;
                }
                else
                {
                    y--;
                }
            }

            return lcs;
        }

        /// <summary>
        /// 找到两个列表的最长公共子序列（LCS）- 基于原文内容
        /// 只有原文内容相同且唯一的段落才能作为锚点
        /// </summary>
        private static List<string> FindLongestCommonSubsequenceByContent(List<string> initialSnapshot, List<string> finalSnapshot)
        {
            // 首先检查每个原文内容在各自快照中的唯一性
            // 如果某个原文内容在快照中出现多次，则不能作为锚点
            var initialContentCount = new Dictionary<string, int>();
            var finalContentCount = new Dictionary<string, int>();
            
            foreach (var paraName in initialSnapshot)
            {
                string content = GetParagraphContent(paraName);
                if (!string.IsNullOrEmpty(content))
                {
                    int count = initialContentCount.ContainsKey(content) ? initialContentCount[content] : 0;
                    initialContentCount[content] = count + 1;
                }
            }
            
            foreach (var paraName in finalSnapshot)
            {
                string content = GetParagraphContent(paraName);
                if (!string.IsNullOrEmpty(content))
                {
                    int count = finalContentCount.ContainsKey(content) ? finalContentCount[content] : 0;
                    finalContentCount[content] = count + 1;
                }
            }
            
            // 构建基于原文内容的列表（只包含唯一的原文内容）
            var initialContentList = new List<string>();
            var finalContentList = new List<string>();
            var initialNameToContentIndex = new Dictionary<string, int>(); // 名称 -> 在initialContentList中的索引
            
            foreach (var paraName in initialSnapshot)
            {
                string content = GetParagraphContent(paraName);
                if (!string.IsNullOrEmpty(content) && 
                    initialContentCount[content] == 1 && 
                    finalContentCount.ContainsKey(content) && 
                    finalContentCount[content] == 1)
                {
                    // 原文内容唯一且在两个快照中都存在
                    initialContentList.Add(content);
                    initialNameToContentIndex[paraName] = initialContentList.Count - 1;
                }
                else
                {
                    // 原文内容不唯一或不存在于最终快照中，使用特殊标记（确保不匹配）
                    initialContentList.Add($"__NON_UNIQUE__{paraName}");
                    initialNameToContentIndex[paraName] = initialContentList.Count - 1;
                }
            }
            
            foreach (var paraName in finalSnapshot)
            {
                string content = GetParagraphContent(paraName);
                if (!string.IsNullOrEmpty(content) && 
                    finalContentCount[content] == 1 && 
                    initialContentCount.ContainsKey(content) && 
                    initialContentCount[content] == 1)
                {
                    // 原文内容唯一且在两个快照中都存在
                    finalContentList.Add(content);
                }
                else
                {
                    // 原文内容不唯一或不存在于初始快照中，使用特殊标记（确保不匹配）
                    finalContentList.Add($"__NON_UNIQUE__{paraName}");
                }
            }
            
            // 基于原文内容计算LCS
            var contentLcs = FindLongestCommonSubsequence(initialContentList, finalContentList);
            
            // 将内容LCS转换回段落名称
            var nameLcs = new List<string>();
            int initialIdx = 0;
            int finalIdx = 0;
            int contentLcsIdx = 0;
            
            while (initialIdx < initialSnapshot.Count && finalIdx < finalSnapshot.Count && contentLcsIdx < contentLcs.Count)
            {
                string initialName = initialSnapshot[initialIdx];
                string finalName = finalSnapshot[finalIdx];
                string initialContent = GetParagraphContent(initialName);
                string finalContent = GetParagraphContent(finalName);
                string lcsContent = contentLcs[contentLcsIdx];
                
                // 检查是否匹配LCS（原文内容相同且唯一）
                int initialCount = initialContentCount.ContainsKey(initialContent) ? initialContentCount[initialContent] : 0;
                int finalCountForInitial = finalContentCount.ContainsKey(initialContent) ? finalContentCount[initialContent] : 0;
                bool initialMatches = !string.IsNullOrEmpty(initialContent) && 
                                     initialContent == lcsContent &&
                                     initialCount == 1 &&
                                     finalCountForInitial == 1;
                
                int finalCount = finalContentCount.ContainsKey(finalContent) ? finalContentCount[finalContent] : 0;
                int initialCountForFinal = initialContentCount.ContainsKey(finalContent) ? initialContentCount[finalContent] : 0;
                bool finalMatches = !string.IsNullOrEmpty(finalContent) && 
                                   finalContent == lcsContent &&
                                   finalCount == 1 &&
                                   initialCountForFinal == 1;
                
                if (initialMatches && finalMatches)
                {
                    // 两个快照都匹配LCS，且原文内容唯一
                    nameLcs.Add(initialName); // 使用初始快照的名称
                    initialIdx++;
                    finalIdx++;
                    contentLcsIdx++;
                }
                else if (initialMatches && !finalMatches)
                {
                    finalIdx++;
                }
                else if (!initialMatches && finalMatches)
                {
                    initialIdx++;
                }
                else
                {
                    if (initialIdx < initialSnapshot.Count) initialIdx++;
                    if (finalIdx < finalSnapshot.Count) finalIdx++;
                }
            }
            
            return nameLcs;
        }

        /// <summary>
        /// 查找锚点段落（在初始和最终快照中都存在且位置相同的段落）
        /// </summary>
        private static List<string> FindAnchorPoints()
        {
            var anchors = new List<string>();

            if (_snapshots.Count < 2)
                return anchors;

            var initialSnapshot = _snapshots[0];
            var finalSnapshot = _snapshots[_snapshots.Count - 1];

            foreach (var name in initialSnapshot)
            {
                if (finalSnapshot.Contains(name) &&
                    HasParagraphName(name) &&
                    initialSnapshot.IndexOf(name) == finalSnapshot.IndexOf(name))
                {
                    anchors.Add(name);
                }
            }

            return anchors;
        }

        /// <summary>
        /// 检查段落名称是否存在
        /// </summary>
        private static bool HasParagraphName(string paragraphName)
        {
            foreach (var mapping in _paragraphNameMapping)
            {
                if (mapping.Count >= 1 && mapping[0] == paragraphName)
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// 获取命名映射表的字符串表示
        /// 格式：[['AAA','哈希码','内容'],['AAB','哈希码','内容']]
        /// </summary>
        private static string GetNameMappingString()
        {
            var entries = _paragraphNameMapping.Select(mapping =>
            {
                if (mapping.Count >= 3)
                {
                    string name = mapping[0];
                    string hash = mapping[1];
                    string content = mapping[2] ?? "";
                    string contentPreview = content.Substring(0, Math.Min(30, content.Length));
                    return $"['{name}','{hash.Substring(0, Math.Min(8, hash.Length))}...','{contentPreview}...']";
                }
                return "[]";
            });
            return "[" + string.Join(",", entries) + "]";
        }

        /// <summary>
        /// 获取快照历史的字符串表示
        /// </summary>
        private static string GetSnapshotsString()
        {
            var snapshotStrings = _snapshots.Select(snapshot =>
                "[" + string.Join(",", snapshot.Select(name => $"\"{name}\"").ToArray()) + "]");
            return "[" + string.Join(",", snapshotStrings) + "]";
        }

        /// <summary>
        /// 清除段落跟踪数据
        /// </summary>
        private static void ClearParagraphTracking()
        {
            _paragraphNameMapping.Clear();
            _snapshots.Clear();
            _nextNameIndex = 0;
            _oldSnapshotIndex = 0; // 重置old_snapshot指针到最初快照
            _displayHunk.Clear(); // 清除display_hunk
        }

        /// <summary>
        /// 设置old_snapshot指针指向指定的快照索引
        /// </summary>
        /// <param name="snapshotIndex">快照索引（从0开始）</param>
        /// <returns>是否设置成功</returns>
        public static bool SetOldSnapshotIndex(int snapshotIndex)
        {
            if (snapshotIndex < 0 || snapshotIndex >= _snapshots.Count)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪] 错误：无效的快照索引 {snapshotIndex}，快照总数: {_snapshots.Count}");
                return false;
            }
            
            _oldSnapshotIndex = snapshotIndex;
            System.Diagnostics.Debug.WriteLine($"[段落跟踪] old_snapshot指针已设置为索引 {snapshotIndex}");
            return true;
        }

        /// <summary>
        /// 根据hunk更新display_hunk
        /// display_hunk存储全量hunk（包含所有段落，不需要header）
        /// </summary>
        /// <param name="hunk">全量hunk列表（只包含一个全量hunk）</param>
        public static void UpdateDisplayHunk(List<(string header, List<string> lines)> hunk)
        {
            // 如果hunk为空（没有改动），display_hunk保持不变
            if (hunk == null || hunk.Count == 0)
            {
                System.Diagnostics.Debug.WriteLine("[段落跟踪] hunk为空，display_hunk保持不变");
                return;
            }

            // 直接存储全量hunk列表（包含所有段落）
            _displayHunk = new List<(string header, List<string> lines)>();
            foreach (var (header, lines) in hunk)
            {
                // 创建hunk的副本
                _displayHunk.Add((header, new List<string>(lines)));
            }
            
            System.Diagnostics.Debug.WriteLine($"[段落跟踪] display_hunk已更新，包含 {_displayHunk.Count} 个全量hunk");
            foreach (var (header, lines) in _displayHunk)
            {
                System.Diagnostics.Debug.WriteLine($"[段落跟踪] display_hunk - header: '{header}', 包含 {lines.Count} 行（全量段落）");
            }
        }

        /// <summary>
        /// 从段落内容中去除接受和丢弃按钮文本
        /// </summary>
        private static string RemoveButtonTextFromContent(Word.Document doc, Word.Range paraRange, string content)
        {
            if (string.IsNullOrEmpty(content))
                return content;

            try
            {
                // 检查段落范围中是否包含书签（以 "operation_" 开头的书签）
                foreach (Word.Bookmark bookmark in doc.Bookmarks)
                {
                    if (bookmark.Name.StartsWith("operation_"))
                    {
                        Word.Range bookmarkRange = bookmark.Range;
                        // 检查书签是否在段落范围内
                        if (bookmarkRange.Start >= paraRange.Start && bookmarkRange.End <= paraRange.End)
                        {
                            // 获取书签在段落中的相对位置
                            int bookmarkStartInPara = bookmarkRange.Start - paraRange.Start;
                            int bookmarkEndInPara = bookmarkRange.End - paraRange.Start;
                            
                            // 确保索引在内容范围内
                            if (bookmarkStartInPara >= 0 && bookmarkEndInPara <= content.Length)
                            {
                                // 从内容中去除书签范围内的文本
                                content = content.Substring(0, bookmarkStartInPara) + content.Substring(bookmarkEndInPara);
                                // 只处理第一个匹配的书签（每个段落通常只有一个按钮组）
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // 如果书签检查失败，尝试通过正则表达式匹配 " 接受  丢弃 " 模式
                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 书签检查失败，尝试正则匹配: {ex.Message}");
                // 使用正则表达式匹配 " 接受 " 和 " 丢弃 " 连在一起的模式
                System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"\s*接受\s+丢弃\s*");
                content = regex.Replace(content, "");
            }

            return content;
        }

        /// <summary>
        /// 处理按钮点击后的操作：更新display_hunk并生成新快照
        /// 第一步：根据操作范围修改display_hunk中对应的操作块
        /// 第二步：根据display_hunk生成新的快照
        /// 第三步：更新old_snapshot指针
        /// </summary>
        /// <param name="operationType">操作类型：replace、delete、insert</param>
        /// <param name="originalRange">原始文本范围（用于替换和删除）</param>
        /// <param name="newRange">新文本范围（用于插入和替换）</param>
        /// <param name="action">操作：accept 或 undo</param>
        public static void HandleButtonClick(string operationType, Word.Range originalRange = null, Word.Range newRange = null, string action = "accept")
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 开始处理操作：类型={operationType}, 动作={action}");

                // 第一步：根据操作范围找到对应的段落名称，然后修改display_hunk
                var affectedParagraphNames = new HashSet<string>();

                // 从originalRange和newRange中提取段落名称
                // Range可能包含多个段落，需要按段落拆分并分别匹配
                if (originalRange != null)
                {
                    Word.Document doc = originalRange.Document;
                    int rangeStart = originalRange.Start;
                    int rangeEnd = originalRange.End;
                    
                    // 遍历文档中的所有段落，找到在Range范围内的段落
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        var para = doc.Paragraphs[i];
                        Word.Range paraRange = para.Range;
                        
                        // 检查段落是否与Range有重叠
                        if (paraRange.Start < rangeEnd && paraRange.End > rangeStart)
                        {
                            string paraContent = paraRange.Text?.TrimEnd('\r', '\n') ?? "";
                            // 去除按钮文本（如果有）
                            paraContent = RemoveButtonTextFromContent(doc, paraRange, paraContent);
                            
                            // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                            string paraName = FindParagraphNameByContent(paraContent);
                            if (paraName != null && !affectedParagraphNames.Contains(paraName))
                            {
                                affectedParagraphNames.Add(paraName);
                                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 从originalRange找到段落: {paraName}");
                            }
                        }
                    }
                }

                if (newRange != null)
                {
                    Word.Document doc = newRange.Document;
                    int rangeStart = newRange.Start;
                    int rangeEnd = newRange.End;
                    
                    // 遍历文档中的所有段落，找到在Range范围内的段落
                    for (int i = 1; i <= doc.Paragraphs.Count; i++)
                    {
                        var para = doc.Paragraphs[i];
                        Word.Range paraRange = para.Range;
                        
                        // 检查段落是否与Range有重叠
                        if (paraRange.Start < rangeEnd && paraRange.End > rangeStart)
                        {
                            string paraContent = paraRange.Text?.TrimEnd('\r', '\n') ?? "";
                            // 去除按钮文本（如果有）
                            paraContent = RemoveButtonTextFromContent(doc, paraRange, paraContent);
                            
                            // 通过内容匹配找到对应的段落名称（使用哈希码验证）
                            string paraName = FindParagraphNameByContent(paraContent);
                            if (paraName != null && !affectedParagraphNames.Contains(paraName))
                            {
                                affectedParagraphNames.Add(paraName);
                                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 从newRange找到段落: {paraName}");
                            }
                        }
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 找到受影响的段落名称: {string.Join(",", affectedParagraphNames)}");

                // 第一步：在display_hunk中找到对应的操作块并修改
                if (_displayHunk.Count > 0)
                {
                    var hunkLines = _displayHunk[0].lines;
                    
                    // 识别操作块：找到包含受影响段落名称的操作块
                    int operationBlockStart = -1;
                    int operationBlockEnd = -1;
                    
                    for (int i = 0; i < hunkLines.Count; i++)
                    {
                        string line = hunkLines[i];
                        string paraName = null;
                        
                        // 提取段落名称
                        if (line.StartsWith("-"))
                        {
                            paraName = line.Substring(1);
                        }
                        else if (line.StartsWith("+"))
                        {
                            paraName = line.Substring(1);
                        }
                        else if (!string.IsNullOrWhiteSpace(line))
                        {
                            paraName = line;
                        }
                        
                        // 检查这个段落是否在受影响范围内
                        if (paraName != null && affectedParagraphNames.Contains(paraName))
                        {
                            // 找到操作块的起始位置（向前找到第一个改动行或操作块开始）
                            if (operationBlockStart == -1)
                            {
                                operationBlockStart = i;
                                // 向前查找，找到操作块的真正开始（连续的改动行的开始）
                                for (int j = i - 1; j >= 0; j--)
                                {
                                    string prevLine = hunkLines[j];
                                    if (prevLine.StartsWith("-") || prevLine.StartsWith("+"))
                                    {
                                        operationBlockStart = j;
                                    }
                                    else if (!string.IsNullOrWhiteSpace(prevLine))
                                    {
                                        // 遇到未改动的行，停止
                                        break;
                                    }
                                }
                            }
                            
                            // 更新操作块的结束位置
                            operationBlockEnd = i;
                        }
                    }
                    
                    // 向后查找，找到操作块的真正结束（连续的改动行的结束）
                    if (operationBlockEnd >= 0)
                    {
                        for (int j = operationBlockEnd + 1; j < hunkLines.Count; j++)
                        {
                            string nextLine = hunkLines[j];
                            if (nextLine.StartsWith("-") || nextLine.StartsWith("+"))
                            {
                                operationBlockEnd = j;
                            }
                            else if (!string.IsNullOrWhiteSpace(nextLine))
                            {
                                // 遇到未改动的行，停止
                                break;
                            }
                        }
                    }
                    
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 找到操作块范围: {operationBlockStart} - {operationBlockEnd}");
                    
                    // 修改操作块内的行
                    var modifiedLines = new List<string>();
                    
                    for (int i = 0; i < hunkLines.Count; i++)
                    {
                        string line = hunkLines[i];
                        
                        // 检查是否在操作块范围内
                        if (i >= operationBlockStart && i <= operationBlockEnd && operationBlockStart >= 0)
                        {
                            // 在操作块范围内
                            if (action == "accept")
                            {
                                // 接受：只保留+号的行（去掉+号前缀），删除-号的行和未改动的行
                                if (line.StartsWith("+"))
                                {
                                    // 保留+号的行，去掉+号前缀
                                    modifiedLines.Add(line.Substring(1));
                                }
                                // -号的行和未改动的行都删除（在操作块范围内）
                            }
                            else // undo
                            {
                                // 丢弃：只保留-号的行（去掉-号前缀），删除+号的行和未改动的行
                                if (line.StartsWith("-"))
                                {
                                    // 保留-号的行，去掉-号前缀
                                    modifiedLines.Add(line.Substring(1));
                                }
                                // +号的行和未改动的行都删除（在操作块范围内）
                            }
                        }
                        else
                        {
                            // 不在操作块范围内的行，直接保留
                            modifiedLines.Add(line);
                        }
                    }
                    
                    // 更新display_hunk
                    _displayHunk[0] = ("", modifiedLines);
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] display_hunk已更新，包含 {modifiedLines.Count} 行");
                }

                // 第二步：根据display_hunk生成新的快照（未改动行 + -号行）
                // 取display_hunk中-号的行和未改动的行组成新的快照
                if (_displayHunk.Count > 0)
                {
                    var hunkLines = _displayHunk[0].lines;
                    var snapshotFromMinus = new List<string>();

                    foreach (var line in hunkLines)
                    {
                        if (line.StartsWith("-"))
                        {
                            // -号的行：去掉-号前缀，添加到新快照
                            snapshotFromMinus.Add(line.Substring(1));
                        }
                        else if (!line.StartsWith("+") && !string.IsNullOrWhiteSpace(line))
                        {
                            // 未改动的行：直接添加到新快照
                            snapshotFromMinus.Add(line);
                        }
                        // +号的行不添加到这个快照
                    }

                    // 插入到历史快照中的最后
                    _snapshots.Add(snapshotFromMinus);
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 快照1（未改动行+-号行）已创建: {string.Join(",", snapshotFromMinus)}");

                    // 第三步：把old_snapshot的指针指到我们刚添加的快照中
                    _oldSnapshotIndex = _snapshots.Count - 1;
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] old_snapshot指针已更新到索引 {_oldSnapshotIndex}");

                    // 第四步：根据display_hunk生成新的快照（未改动行 + +号行）
                    // 取display_hunk中+号的行和未改动的行组成新的快照
                    var snapshotFromPlus = new List<string>();

                    foreach (var line in hunkLines)
                    {
                        if (line.StartsWith("+"))
                        {
                            // +号的行：去掉+号前缀，添加到新快照
                            snapshotFromPlus.Add(line.Substring(1));
                        }
                        else if (!line.StartsWith("-") && !string.IsNullOrWhiteSpace(line))
                        {
                            // 未改动的行：直接添加到新快照
                            snapshotFromPlus.Add(line);
                        }
                        // -号的行不添加到这个快照
                    }

                    // 插入到历史快照中的最后
                    _snapshots.Add(snapshotFromPlus);
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 快照2（未改动行++号行）已创建: {string.Join(",", snapshotFromPlus)}");
                    System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 快照历史: {GetSnapshotsString()}");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[按钮点击处理] 处理失败: {ex.Message}");
            }
        }

        #endregion
    }
}
