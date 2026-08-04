using System;
using System.Linq;
using System.Web.Script.Serialization;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 文档处理器类，用于在Word文档中执行各种操作
    /// </summary>
    public class DocumentProcessor
    {
        private Word.Application _application;

        // 存储操作信息的静态字典
        private static Dictionary<string, object> _operationInfos = new Dictionary<string, object>();

        /// <summary>
        /// 获取操作信息字典（用于外部访问）
        /// </summary>
        public static Dictionary<string, object> OperationInfos => _operationInfos;

        /// <summary>
        /// 初始化文档处理器
        /// </summary>
        /// <param name="application">Word应用程序实例</param>
        public DocumentProcessor(Word.Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// 将连续的数字列表转换为区间列表
        /// </summary>
        /// <param name="positions">已排序的数字列表</param>
        /// <returns>区间列表</returns>
        private List<(int Start, int End)> ConvertPositionsToRanges(List<int> positions)
        {
            var ranges = new List<(int Start, int End)>();
            if (positions == null || positions.Count == 0)
            {
                return ranges;
            }

            int start = positions[0];
            int current = positions[0];

            for (int i = 1; i < positions.Count; i++)
            {
                if (positions[i] == current + 1)
                {
                    // 连续，继续扩展
                    current = positions[i];
                }
                else
                {
                    // 不连续，保存当前区间，开始新区间
                    ranges.Add((start, current));
                    start = positions[i];
                    current = positions[i];
                }
            }

            // 添加最后一个区间
            ranges.Add((start, current));

            return ranges;
        }

        /// <summary>
        /// 根据对消位置删除文档中的文字
        /// </summary>
        /// <param name="eliminationPositions">需要删除的位置列表（已排序）</param>
        public void ApplyDocumentElimination(List<int> eliminationPositions)
        {
            if (eliminationPositions == null || eliminationPositions.Count == 0)
            {
                return;
            }

            try
            {
                Word.Document document = _application.ActiveDocument;

                // 将连续的对消位置合并为区间，提高删除效率
                var eliminationRanges = ConvertPositionsToRanges(eliminationPositions);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 对消区间: {string.Join(",", eliminationRanges.Select(r => $"[{r.Start},{r.End}]"))}");

                // 从后往前删除区间，避免位置变化影响后续删除
                for (int i = eliminationRanges.Count - 1; i >= 0; i--)
                {
                    var range = eliminationRanges[i];

                    try
                    {
                        // 获取指定区间的范围（注意：Word.Range的结束位置是exclusive的，所以要+1）
                        Word.Range deleteRange = document.Range(range.Start, range.End + 1);
                        if (deleteRange != null && deleteRange.Text != null)
                        {
                            // 保存原始文本用于调试输出
                            string originalText = deleteRange.Text;
                            // 删除这个区间的内容
                            deleteRange.Text = "";
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] 文档对消：删除了区间 [{range.Start},{range.End}] 的文字 '{originalText.Replace("\r", "\\r").Replace("\n", "\\n")}'");
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 删除区间 [{range.Start},{range.End}] 时出错: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] 文档对消完成，共删除了 {eliminationPositions.Count} 个字符，合并为 {eliminationRanges.Count} 个区间");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 文档对消处理出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 根据JSON指令在文档中执行操作
        /// JSON格式: {"type":"insert","detail":{"pre":"xxx","insert":"xxx","post":"xxx"}}
        /// </summary>
        /// <param name="jsonInput">JSON格式的操作指令</param>
        /// <param name="insertedRange">输出参数，返回插入内容的范围</param>
        /// <returns>操作是否成功</returns>
        public bool ProcessContent(string jsonInput, out (int Start, int End)? insertedRange)
        {
            insertedRange = null; // 默认没有插入范围

            try
            {
                // 解析JSON
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> root = serializer.Deserialize<Dictionary<string, object>>(jsonInput);

                // 验证JSON格式
                string operationType = root["type"]?.ToString();
                if (string.IsNullOrEmpty(operationType) || (operationType != "insert" && operationType != "replace" && operationType != "delete"))
                {
                    throw new ArgumentException("无效的JSON格式：type必须为'insert'、'replace'或'delete'");
                }

                if (!root.ContainsKey("detail"))
                {
                    throw new ArgumentException("无效的JSON格式：缺少detail字段");
                }

                Dictionary<string, object> detail = root["detail"] as Dictionary<string, object>;
                if (detail == null)
                {
                    throw new ArgumentException("无效的JSON格式：detail字段格式错误");
                }

                if (operationType == "insert")
                {
                    // 处理插入操作
                    string pre = detail.ContainsKey("pre") ? detail["pre"]?.ToString() : null;
                    string insert = detail.ContainsKey("insert") ? detail["insert"]?.ToString() : null;
                    string post = detail.ContainsKey("post") ? detail["post"]?.ToString() : null;

                    if (string.IsNullOrEmpty(pre) || string.IsNullOrEmpty(insert) || post == null)
                    {
                        throw new ArgumentException("无效的JSON格式：insert操作的detail字段缺少pre、insert或post字段");
                    }

                    // 执行插入操作
                    return InsertContentBetweenContext(pre, insert, post, out insertedRange);
                }
                else if (operationType == "replace")
                {
                    // 处理替换操作
                    object originalObj = detail.ContainsKey("original") ? detail["original"] : null;
                    string newText = detail.ContainsKey("new") ? detail["new"]?.ToString() : null;

                    if (originalObj == null || newText == null)
                    {
                        throw new ArgumentException("无效的JSON格式：replace操作的detail字段缺少original或new字段");
                    }

                    // 执行替换操作
                    return ReplaceText(originalObj, newText);
                }
                else if (operationType == "delete")
                {
                    // 处理删除操作
                    object deletionObj = detail.ContainsKey("deletion") ? detail["deletion"] : null;

                    if (deletionObj == null)
                    {
                        throw new ArgumentException("无效的JSON格式：delete操作的detail字段缺少deletion字段");
                    }

                    // 执行删除操作
                    return DeleteText(deletionObj);
                }
                else
                {
                    throw new ArgumentException($"不支持的操作类型：{operationType}");
                }
            }
            catch (ArgumentException)
            {
                throw;
            }
            catch (InvalidOperationException)
            {
                throw;
            }
            catch (Exception ex)
            {
                throw new Exception($"操作执行失败：{ex.Message}", ex);
            }
        }

        /// <summary>
        /// 在文档中找到上下文并插入内容
        /// </summary>
        /// <param name="pre">上文内容</param>
        /// <param name="insert">要插入的内容</param>
        /// <param name="post">下文内容</param>
        /// <returns>插入是否成功</returns>
        private bool InsertContentBetweenContext(string pre, string insert, string post, out (int Start, int End)? insertedRange)
        {
            insertedRange = null;

            // 检查是否有活动文档
            if (_application.ActiveDocument == null)
            {
                throw new InvalidOperationException("没有活动的Word文档");
            }

            Word.Document document = _application.ActiveDocument;
            Word.Range documentRange = document.Content;

            // 搜索上文内容
            Word.Range preRange = FindText(document.Content, pre);
            if (preRange == null)
            {
                throw new InvalidOperationException($"在文档中找不到上文内容：{pre}");
            }

            Word.Range insertRange;
            if (string.IsNullOrEmpty(post))
            {
                // 如果post为空，在pre文本后直接插入
                insertRange = document.Range(preRange.End, preRange.End);
            }
            else
            {
                // 搜索下文内容
                Word.Range postRange = FindText(document.Content, post);
                if (postRange == null)
                {
                    throw new InvalidOperationException($"在文档中找不到下文内容：{post}");
                }

                // 确保上文在下文之前
                if (preRange.End >= postRange.Start)
                {
                    throw new InvalidOperationException("上文内容的位置不能在下文内容之后");
                }

                // 在上文后插入内容（使用collapsed range确保是插入而不是替换）
                insertRange = document.Range(preRange.End, preRange.End);
            }

            // 记录插入起始位置
            int insertStartPosition = insertRange.Start;

            // 只插入文本，不应用任何格式，不记录历史，不添加按钮
            insertRange.Text = insert;

            // 打印插入后的完整文档内容
            string fullContent = insertRange.Document.Content.Text;
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入后完整文档内容 ({fullContent.Length}字符): '{fullContent}'");

            // 插入操作后，Range会扩展为插入内容的范围
            insertedRange = (insertStartPosition, insertStartPosition + insert.Length - 1);

            return true;
        }

        /// <summary>
        /// 替换文档中的文本
        /// </summary>
        /// <param name="original">要替换的原始文本（字符串或包含start/end的对象）</param>
        /// <param name="newText">新的文本内容</param>
        /// <returns>替换是否成功</returns>
        private bool ReplaceText(object original, string newText)
        {
            // 检查是否有活动文档
            if (_application.ActiveDocument == null)
            {
                throw new InvalidOperationException("没有活动的Word文档");
            }

            Word.Document document = _application.ActiveDocument;

            Word.Range originalRange = null;

            // 处理original参数的不同格式
            if (original is string originalText)
            {
                // 原始格式：直接搜索文本
                originalRange = FindText(document.Content, originalText);
                if (originalRange == null)
                {
                    throw new InvalidOperationException($"在文档中找不到要替换的文本：{originalText}");
                }
            }
            else if (original is Dictionary<string, object> originalDict)
            {
                // 新格式：通过start和end字符查找
                string startText = originalDict.ContainsKey("start") ? originalDict["start"]?.ToString() : null;
                string endText = originalDict.ContainsKey("end") ? originalDict["end"]?.ToString() : null;

                if (string.IsNullOrEmpty(startText) || string.IsNullOrEmpty(endText))
                {
                    throw new ArgumentException("无效的original格式：缺少start或end字段");
                }

                originalRange = FindTextByStartEnd(document.Content, startText, endText);
                if (originalRange == null)
                {
                    throw new InvalidOperationException($"在文档中找不到要替换的文本（从'{startText}'到'{endText}'）");
                }
            }
            else
            {
                throw new ArgumentException("无效的original格式：必须是字符串或包含start/end的对象");
            }

            // 只在原有文本后面插入新文本，不删除原有文本
            // 在原有文本后面插入新文本
            Word.Range insertRange = document.Range(originalRange.End, originalRange.End);
            insertRange.Text = newText;

            // 打印插入后的完整文档内容
            string fullContent = document.Content.Text;
            System.Diagnostics.Debug.WriteLine($"[DEBUG] 插入后完整文档内容 ({fullContent.Length}字符): '{fullContent}'");

            // 注意：不应用任何格式，不记录历史，不添加按钮
            // 所有的格式、历史记录和按钮都将在第9步统一处理

            return true;
        }

        /// <summary>
        /// 删除文档中的文本（添加删除线）
        /// </summary>
        /// <param name="deletion">要删除的文本（字符串或包含start/end的对象）</param>
        /// <returns>删除是否成功</returns>
        private bool DeleteText(object deletion)
        {
            // 检查是否有活动文档
            if (_application.ActiveDocument == null)
            {
                throw new InvalidOperationException("没有活动的Word文档");
            }

            Word.Document document = _application.ActiveDocument;

            Word.Range deletionRange = null;

            // 处理deletion参数的不同格式
            if (deletion is string deletionText)
            {
                // 原始格式：直接搜索文本
                deletionRange = FindText(document.Content, deletionText);
                if (deletionRange == null)
                {
                    throw new InvalidOperationException($"在文档中找不到要删除的文本：{deletionText}");
                }
            }
            else if (deletion is Dictionary<string, object> deletionDict)
            {
                // 新格式：通过start和end字符查找
                string startText = deletionDict.ContainsKey("start") ? deletionDict["start"]?.ToString() : null;
                string endText = deletionDict.ContainsKey("end") ? deletionDict["end"]?.ToString() : null;

                if (string.IsNullOrEmpty(startText) || string.IsNullOrEmpty(endText))
                {
                    throw new ArgumentException("无效的deletion格式：缺少start或end字段");
                }

                deletionRange = FindTextByStartEnd(document.Content, startText, endText);
                if (deletionRange == null)
                {
                    throw new InvalidOperationException($"在文档中找不到要删除的文本（从'{startText}'到'{endText}'）");
                }
            }
            else
            {
                throw new ArgumentException("无效的deletion格式：必须是字符串或包含start/end的对象");
            }

            // 简单删除文本，不应用任何格式，不记录历史，不添加按钮
            deletionRange.Text = "";

            return true;
        }

        /// <summary>
        /// 在文档中插入内联的接受和丢弃按钮
        /// </summary>
        /// <param name="range">文本范围</param>
        /// <param name="operationType">操作类型</param>
        /// <param name="originalRange">原始文本范围（用于替换和删除）</param>
        /// <param name="newRange">新文本范围（用于插入和替换）</param>
        /// <param name="originalSentenceNames">原始句子名称列表（句子级别操作使用）</param>
        /// <param name="newSentenceNames">新句子名称列表（句子级别操作使用）</param>
        public void InsertInlineButtons(Word.Range range, string operationType, Word.Range originalRange = null, Word.Range newRange = null, List<string> originalSentenceNames = null, List<string> newSentenceNames = null)
        {
            try
            {
                Word.Document document = _application.ActiveDocument;

                // 在操作文本后面插入按钮文本
                Word.Range buttonRange = document.Range(range.End, range.End);

                // 生成唯一的操作ID（只包含基本信息，不包含时间戳以避免匹配问题）
                string operationId = $"operation_{operationType}_{range.Start}_{range.End}";

                // 如果已存在相同ID的书签，添加时间戳后缀
                string finalOperationId = operationId;
                int counter = 1;
                while (document.Bookmarks.Exists(finalOperationId))
                {
                    finalOperationId = $"{operationId}_{counter}";
                    counter++;
                }

                // 存储操作信息到静态字典中，供后续点击处理使用
                var operationInfo = new
                {
                    OperationType = operationType,
                    OperationRange = range,
                    OriginalRange = originalRange,
                    NewRange = newRange,
                    OperationId = finalOperationId,
                    OriginalSentenceNames = originalSentenceNames, // 句子级别操作使用
                    NewSentenceNames = newSentenceNames // 句子级别操作使用
                };

                // 将操作信息存储到静态集合中
                StoreOperationInfo(finalOperationId, operationInfo);

                // 插入接受按钮
                Word.Range acceptButtonRange = document.Range(range.End, range.End);
                acceptButtonRange.Text = " 接受 ";
                acceptButtonRange.Font.Color = Word.WdColor.wdColorBlack; // 黑色文字
                acceptButtonRange.Font.Size = 9; // 小字体
                acceptButtonRange.Font.Bold = 1; // 加粗
                acceptButtonRange.Shading.BackgroundPatternColor = ConfigManager.DocumentColorHelper.AcceptButtonBackground;

                // 插入丢弃按钮
                Word.Range rejectButtonRange = document.Range(acceptButtonRange.End, acceptButtonRange.End);
                rejectButtonRange.Text = " 丢弃 ";
                rejectButtonRange.Font.Color = Word.WdColor.wdColorBlack; // 黑色文字
                rejectButtonRange.Font.Size = 9; // 小字体
                rejectButtonRange.Font.Bold = 1; // 加粗
                rejectButtonRange.Shading.BackgroundPatternColor = ConfigManager.DocumentColorHelper.RejectButtonBackground;

                // 合并按钮范围作为整体书签
                Word.Range combinedButtonRange = document.Range(acceptButtonRange.Start, rejectButtonRange.End);
                document.Bookmarks.Add(finalOperationId, combinedButtonRange);
            }
            catch (Exception)
            {
                // 忽略插入按钮时的异常
            }
        }

        /// <summary>
        /// 获取Word范围在屏幕上的位置
        /// </summary>
        /// <param name="range">Word范围</param>
        /// <returns>屏幕坐标</returns>
        private Point GetRangeScreenPosition(Word.Range range)
        {
            try
            {
                // 获取范围的边界矩形信息
                _application.ActiveWindow.GetPoint(out int left, out int top, out int width, out int height, range);

                // 转换为屏幕坐标
                Point wordPoint = new Point(left, top);

                // 获取Word窗口的句柄
                IntPtr wordHandle = (IntPtr)_application.ActiveWindow.Hwnd;

                // 安全地将Word坐标转换为屏幕坐标
                Control wordControl = Control.FromHandle(wordHandle);
                if (wordControl != null)
                {
                    Point screenPoint = wordControl.PointToScreen(wordPoint);
                    return screenPoint;
                }
                else
                {
                    // 如果无法获取控件，使用估算的屏幕位置
                    // Word窗口通常在屏幕中央，添加一些偏移
                    return new Point(left + 50, top + 50);
                }
            }
            catch (Exception)
            {
                // 提供一个更合理的默认位置
                return new Point(300, 200);
            }
        }

        /// <summary>
        /// 获取只包含文本内容的范围（不包括段落标记和后续空白）
        /// </summary>
        private Word.Range GetContentOnlyRange(Word.Range range)
        {
            try
            {
                if (range == null)
                    return null;

                string rangeText = range.Text ?? "";
                
                // 如果范围末尾有段落标记（\r或\n），排除它
                if (rangeText.EndsWith("\r") || rangeText.EndsWith("\n"))
                {
                    // 创建一个新的范围，排除最后一个字符（段落标记）
                    int endPos = range.End - 1;
                    if (endPos >= range.Start)
                    {
                        return _application.ActiveDocument.Range(range.Start, endPos);
                    }
                }
                
                return range;
            }
            catch (Exception)
            {
                return range; // 如果出错，返回原范围
            }
        }

        /// <summary>
        /// 确认操作 - 根据操作类型执行最终确认
        /// </summary>
        /// <param name="operationType">操作类型</param>
        /// <param name="range">操作范围</param>
        /// <param name="originalRange">原始文本范围（用于替换和删除）</param>
        /// <param name="newRange">新文本范围（用于插入和替换）</param>
        /// <param name="originalSentenceNames">原始句子名称列表（句子级别操作使用）</param>
        /// <param name="newSentenceNames">新句子名称列表（句子级别操作使用）</param>
        public void ConfirmOperation(string operationType, Word.Range range, Word.Range originalRange = null, Word.Range newRange = null, List<string> originalSentenceNames = null, List<string> newSentenceNames = null)
        {
            try
            {
                switch (operationType.ToLower())
                {
                    case "insert":
                        // 插入操作确认：将新文字背景从黄色改为白色（字符级别）
                        if (newRange != null)
                        {
                            Word.Range contentRange = GetContentOnlyRange(newRange);
                            if (contentRange != null && contentRange.Text.Length > 0)
                            {
                                // 遍历字符清除黄色背景
                                for (int i = 1; i <= contentRange.Characters.Count; i++)
                                {
                                    var charRange = contentRange.Characters[i];
                                    if (charRange.Shading.BackgroundPatternColor == ConfigManager.DocumentColorHelper.InsertBackground)
                                    {
                                        charRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                                    }
                                }
                            }
                        }
                        break;

                    case "replace":
                        // 替换操作确认：先清除新文字背景色，再删除有删除线的原文（字符级别）
                        // 注意：必须先清除背景色，因为删除原文后newRange的位置会变化，可能导致Range失效
                        if (newRange != null)
                        {
                            Word.Range contentRange = GetContentOnlyRange(newRange);
                            if (contentRange != null && contentRange.Text.Length > 0)
                            {
                                // 遍历字符清除黄色背景
                                for (int i = 1; i <= contentRange.Characters.Count; i++)
                                {
                                    var charRange = contentRange.Characters[i];
                                    if (charRange.Shading.BackgroundPatternColor == ConfigManager.DocumentColorHelper.InsertBackground)
                                    {
                                        charRange.Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite;
                                    }
                                }
                            }
                        }
                        if (originalRange != null)
                        {
                            originalRange.Text = ""; // 删除有删除线的原文
                        }
                        break;

                    case "delete":
                        // 删除操作确认：删除有删除线的原文
                        if (originalRange != null)
                        {
                            originalRange.Text = ""; // 删除有删除线的原文
                        }
                        break;
                }

                // 确认操作后，删除对应的历史记录（内容已被确认，格式已被清理）
                RemoveHistoricalRecordsByOperation(operationType, originalRange, newRange);

                // 处理display_hunk和快照更新
                // 检测是句子级别还是段落级别：如果SentenceState有快照，就是句子级别
                if (SentenceState.CurrentSnapshot != null && SentenceState.CurrentSnapshot.Count > 0)
                {
                    // 句子级别操作：优先使用存储的句子名称列表
                    SentenceState.HandleButtonClick(operationType, originalRange, newRange, "accept", originalSentenceNames, newSentenceNames);
                }
                else
                {
                    // 段落级别操作
                    DocumentState.HandleButtonClick(operationType, originalRange, newRange, "accept");
                }
            }
            catch (Exception)
            {
                // 忽略确认操作时的异常
            }
        }

        /// <summary>
        /// 丢弃操作 - 撤销之前的操作
        /// </summary>
        /// <param name="operationType">操作类型</param>
        /// <param name="range">操作范围</param>
        /// <param name="originalRange">原始文本范围（用于替换和删除）</param>
        /// <param name="newRange">新文本范围（用于插入和替换）</param>
        /// <param name="originalSentenceNames">原始句子名称列表（句子级别操作使用）</param>
        /// <param name="newSentenceNames">新句子名称列表（句子级别操作使用）</param>
        public void RejectOperation(string operationType, Word.Range range, Word.Range originalRange = null, Word.Range newRange = null, List<string> originalSentenceNames = null, List<string> newSentenceNames = null)
        {
            try
            {
                switch (operationType.ToLower())
                {
                    case "insert":
                        // 插入操作丢弃：完全删除新文字
                        if (newRange != null)
                        {
                            newRange.Text = "";
                        }
                        break;

                    case "replace":
                        // 替换操作丢弃：恢复原文（去掉删除线），完全删除新文字
                        if (originalRange != null)
                        {
                            originalRange.Font.StrikeThrough = 0; // 去掉删除线
                        }
                        if (newRange != null)
                        {
                            newRange.Text = "";
                        }
                        break;

                    case "delete":
                        // 删除操作丢弃：恢复原文（去掉删除线）
                        if (originalRange != null)
                        {
                            originalRange.Font.StrikeThrough = 0; // 去掉删除线
                        }
                        break;
                }

                // 丢弃操作后，删除对应的历史记录（操作被撤销）
                RemoveHistoricalRecordsByOperation(operationType, originalRange, newRange);

                // 处理display_hunk和快照更新
                // 检测是句子级别还是段落级别：如果SentenceState有快照，就是句子级别
                if (SentenceState.CurrentSnapshot != null && SentenceState.CurrentSnapshot.Count > 0)
                {
                    // 句子级别操作：优先使用存储的句子名称列表
                    SentenceState.HandleButtonClick(operationType, originalRange, newRange, "reject", originalSentenceNames, newSentenceNames);
                }
                else
                {
                    // 段落级别操作
                    DocumentState.HandleButtonClick(operationType, originalRange, newRange, "reject");
                }
            }
            catch (Exception)
            {
                // 忽略拒绝操作时的异常
            }
        }

        /// <summary>
        /// 通过开始和结束字符查找文本范围
        /// </summary>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="startText">开始字符</param>
        /// <param name="endText">结束字符</param>
        /// <returns>找到的范围，如果没找到则返回null</returns>
        public Word.Range FindTextByStartEnd(Word.Range searchRange, string startText, string endText)
        {
            // 默认调用考虑历史删除记录的版本，传入空的删除记录列表
            return FindTextByStartEndValid(searchRange, startText, endText, new List<(int Start, int End)>(), false);
        }

        /// <summary>
        /// 在指定范围内搜索所有匹配的文本
        /// </summary>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="searchText">要搜索的文本</param>
        /// <returns>所有找到的范围列表</returns>
        public List<Word.Range> FindAllText(Word.Range searchRange, string searchText)
        {
            var results = new List<Word.Range>();

            if (string.IsNullOrEmpty(searchText))
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - searchText为空，返回空结果");
                return results;
            }

            System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 开始搜索文本: '{searchText}'，搜索范围长度: {searchRange.End - searchRange.Start}");

            try
            {
                // 创建搜索范围的副本，避免Find操作修改原始范围
                Word.Range searchRangeCopy = searchRange.Duplicate;
                System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 创建搜索范围副本成功");

                // 设置查找参数
                Word.Find find = searchRangeCopy.Find;
                find.Text = searchText;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;

                // 循环查找所有匹配
                int findCount = 0;
                int lastMatchStart = -1;
                int lastMatchEnd = -1;

                while (find.Execute())
                {
                    findCount++;
                    int currentMatchStart = searchRangeCopy.Start;
                    int currentMatchEnd = searchRangeCopy.End;

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 找到第{findCount}个匹配: 位置[{currentMatchStart}, {currentMatchEnd}]");

                    // 检查是否是重复的匹配（防止死循环）
                    if (currentMatchStart == lastMatchStart && currentMatchEnd == lastMatchEnd)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 检测到重复匹配，停止搜索以防止死循环");
                        break;
                    }

                    // 保存当前匹配
                    Word.Range matchRange = searchRangeCopy.Duplicate;
                    results.Add(matchRange);

                    // 记录这次匹配的位置
                    lastMatchStart = currentMatchStart;
                    lastMatchEnd = currentMatchEnd;

                    // 移动到下一个位置继续搜索（强制前进至少一个字符）
                    int nextStart = Math.Max(currentMatchEnd, currentMatchStart + 1);
                    searchRangeCopy.Start = nextStart;
                    searchRangeCopy.End = searchRange.End;

                    // 如果已经到达搜索范围的末尾，停止
                    if (searchRangeCopy.Start >= searchRange.End)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 已到达搜索范围末尾，停止搜索");
                        break;
                    }

                    // 防止过度循环
                    if (findCount > 1000)
                    {
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 查找次数过多({findCount})，强制停止以防止死循环");
                        break;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] FindAllText - 搜索完成，共找到{results.Count}个匹配");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 查找所有文本 '{searchText}' 时出错: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 异常类型: {ex.GetType().FullName}");
                if (ex is System.OperationCanceledException)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 检测到OperationCanceledException，这可能与Word操作冲突有关");
                }
            }

            return results;
        }

        /// <summary>
        /// 从多个匹配中选择不在删除区间内的匹配
        /// </summary>
        /// <param name="matches">所有匹配的范围列表</param>
        /// <param name="historicalDeletes">历史删除区间列表</param>
        /// <param name="requireUnique">是否要求只有一个有效匹配</param>
        /// <returns>有效的匹配范围，如果没有或不唯一则返回null</returns>
        public Word.Range SelectValidMatch(List<Word.Range> matches, IReadOnlyList<(int Start, int End)> historicalDeletes, bool requireUnique = false)
        {
            if (matches == null || matches.Count == 0)
            {
                return null;
            }

            var validMatches = new List<Word.Range>();

            // 收集所有不与历史删除记录重叠的匹配
            foreach (var match in matches)
            {
                bool isOverlappingWithDeleted = false;

                foreach (var (deleteStart, deleteEnd) in historicalDeletes)
                {
                    // 检查整个匹配范围是否与历史删除记录重叠，重叠一个字符也不行
                    if (RangesOverlap(match.Start, match.End, deleteStart, deleteEnd))
                    {
                        isOverlappingWithDeleted = true;
                        break;
                    }
                }

                if (!isOverlappingWithDeleted)
                {
                    validMatches.Add(match);
                }
            }

            // 如果需要唯一匹配但找到多个，返回null表示不唯一
            if (requireUnique && validMatches.Count > 1)
            {
                return null; // 表示找到多个有效匹配，特异性不够
            }

            // 如果有有效匹配，返回第一个
            if (validMatches.Count > 0)
            {
                return validMatches[0];
            }

            // 如果所有匹配都在删除区间内，返回null
            return null;
        }

        /// <summary>
        /// 检查有效匹配的数量
        /// </summary>
        /// <param name="matches">所有匹配的范围列表</param>
        /// <param name="historicalDeletes">历史删除区间列表</param>
        /// <returns>有效匹配的数量</returns>
        public int GetValidMatchCount(List<Word.Range> matches, IReadOnlyList<(int Start, int End)> historicalDeletes)
        {
            if (matches == null || matches.Count == 0)
            {
                return 0;
            }

            int validCount = 0;

            foreach (var match in matches)
            {
                bool isOverlappingWithDeleted = false;

                foreach (var (deleteStart, deleteEnd) in historicalDeletes)
                {
                    // 检查整个匹配范围是否与历史删除记录重叠，重叠一个字符也不行
                    if (RangesOverlap(match.Start, match.End, deleteStart, deleteEnd))
                    {
                        isOverlappingWithDeleted = true;
                        break;
                    }
                }

                if (!isOverlappingWithDeleted)
                {
                    validCount++;
                }
            }

            return validCount;
        }

        /// <summary>
        /// 检查两个范围是否重叠
        /// </summary>
        private static bool RangesOverlap(int start1, int end1, int start2, int end2)
        {
            return start1 < end2 && start2 < end1;
        }

        /// <summary>
        /// 通过start和end文本查找范围，选择不在删除区间内的有效匹配
        /// </summary>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="startText">开始文本</param>
        /// <param name="endText">结束文本</param>
        /// <param name="historicalDeletes">历史删除区间列表</param>
        /// <param name="requireUnique">是否要求start和end匹配都唯一</param>
        /// <returns>找到的有效范围，如果没找到或不唯一则返回null</returns>
        public Word.Range FindTextByStartEndValid(Word.Range searchRange, string startText, string endText, IReadOnlyList<(int Start, int End)> historicalDeletes, bool requireUnique = false)
        {
            if (string.IsNullOrEmpty(startText) || string.IsNullOrEmpty(endText))
            {
                return null;
            }

            // 如果start和end文本相同，直接查找并选择有效匹配
            if (startText == endText)
            {
                List<Word.Range> matches = FindAllText(searchRange, startText);
                return SelectValidMatch(matches, historicalDeletes, requireUnique);
            }

            // 在全文中分别查找start和end文本的所有匹配
            Word.Range fullRange = searchRange.Document.Content;

            List<Word.Range> startMatches = FindAllText(fullRange, startText);
            List<Word.Range> endMatches = FindAllText(fullRange, endText);

            // 检查有效匹配数量
            int validStartCount = GetValidMatchCount(startMatches, historicalDeletes);
            int validEndCount = GetValidMatchCount(endMatches, historicalDeletes);

            if (requireUnique)
            {
                if (validStartCount != 1 || validEndCount != 1)
                {
                    return null; // start或end有多个匹配，特异性不够
                }
            }

            // 选择有效的start和end匹配
            Word.Range startRange = SelectValidMatch(startMatches, historicalDeletes, requireUnique);
            Word.Range endRange = SelectValidMatch(endMatches, historicalDeletes, requireUnique);

            if (startRange == null || endRange == null)
            {
                return null;
            }

            // 确定选择范围的起始和结束位置
            int selectionStart = Math.Min(startRange.Start, endRange.Start);
            int selectionEnd = Math.Max(startRange.End, endRange.End);

            // 创建选择范围
            Word.Range resultRange = searchRange.Duplicate;
            resultRange.Start = selectionStart;
            resultRange.End = selectionEnd;

            return resultRange;
        }

        /// <summary>
        /// 检查位置是否在指定范围内（包含结束位置）
        /// </summary>
        private bool PositionInRange(int position, int rangeStart, int rangeEnd)
        {
            return position >= rangeStart && position <= rangeEnd;
        }

        /// <summary>
        /// 在指定范围内搜索文本（只返回第一个匹配）
        /// </summary>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="searchText">要搜索的文本</param>
        /// <returns>找到的范围，如果没找到则返回null</returns>
        public Word.Range FindText(Word.Range searchRange, string searchText)
        {
            if (string.IsNullOrEmpty(searchText))
            {
                return null;
            }

            // 创建搜索范围的副本，避免Find操作修改原始范围
            Word.Range searchRangeCopy = searchRange.Duplicate;

            // 方案1：先尝试精确匹配
            Word.Find find = searchRangeCopy.Find;
            find.Text = searchText;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;

            bool found = find.Execute();

            if (found)
            {
                return searchRangeCopy.Duplicate;
            }

            // 方案2：如果精确匹配失败，尝试智能搜索（忽略删除线内容）
            return SmartFindText(searchRange, searchText);
        }

        /// <summary>
        /// 智能搜索文本，忽略删除线内容进行匹配
        /// </summary>
        private Word.Range SmartFindText(Word.Range searchRange, string searchText)
        {
            try
            {
                if (string.IsNullOrEmpty(searchText))
                    return null;

                Word.Document doc = searchRange.Document;
                string visibleText = "";

                // 收集可见文本和位置映射
                var visibleCharMap = new List<(int docPos, string charText)>();

                for (int i = 1; i <= searchRange.Characters.Count; i++)
                {
                    Word.Range charRange = searchRange.Characters[i];
                    string charText = charRange.Text;

                    // 检查是否应该过滤这个字符（与F_GetDocumentContentTool保持一致）
                    bool shouldFilter = false;

                    Word.WdColor bgColor = charRange.Shading.BackgroundPatternColor;
                    bool hasStrikethrough = charRange.Font.StrikeThrough == -1;

                    Word.WdColor specialGrayBg = ConfigManager.DocumentColorHelper.StrikethroughBackground;
                    Word.WdColor specialYellowBg = ConfigManager.DocumentColorHelper.InsertBackground;
                    Word.WdColor acceptBtnBg = ConfigManager.DocumentColorHelper.AcceptButtonBackground;
                    Word.WdColor rejectBtnBg = ConfigManager.DocumentColorHelper.RejectButtonBackground;

                    // 应用与F_GetDocumentContentTool相同的过滤规则
                    if (hasStrikethrough && bgColor == specialGrayBg)
                    {
                        shouldFilter = true;
                    }
                    else if (bgColor == acceptBtnBg && !hasStrikethrough && charText.Trim().Length > 0)
                    {
                        shouldFilter = true;
                    }
                    else if (bgColor == rejectBtnBg && !hasStrikethrough && charText.Trim().Length > 0)
                    {
                        shouldFilter = true;
                    }

                    if (!shouldFilter)
                    {
                        visibleText += charText;
                        visibleCharMap.Add((charRange.Start, charText));
                    }
                }

                // 在可见文本中查找
                int visibleIndex = visibleText.IndexOf(searchText);
                if (visibleIndex == -1)
                {
                    return null;
                }

                // 将可见文本中的位置映射回文档位置
                int startVisibleIndex = visibleIndex;
                int endVisibleIndex = visibleIndex + searchText.Length - 1;

                if (startVisibleIndex >= visibleCharMap.Count || endVisibleIndex >= visibleCharMap.Count)
                {
                    return null;
                }

                int docStartPos = visibleCharMap[startVisibleIndex].docPos;
                int docEndPos = visibleCharMap[endVisibleIndex].docPos + visibleCharMap[endVisibleIndex].charText.Length;

                Word.Range resultRange = doc.Range(docStartPos, docEndPos);

                return resultRange;
            }
            catch (Exception)
            {
                return null;
            }
        }

        /// <summary>
        /// 存储操作信息以供后续处理
        /// </summary>
        /// <param name="operationId">操作ID</param>
        /// <param name="operationInfo">操作信息</param>
        private void StoreOperationInfo(string operationId, object operationInfo)
        {
            _operationInfos[operationId] = operationInfo;
        }

        /// <summary>
        /// 根据书签名称处理按钮点击操作
        /// </summary>
        /// <param name="bookmarkName">书签名称</param>
        /// <param name="action">操作类型：accept 或 reject</param>
        public void HandleInlineButtonClick(string bookmarkName, string action)
        {
            try
            {
                if (!_operationInfos.ContainsKey(bookmarkName))
                {
                    return;
                }

                var operationInfo = _operationInfos[bookmarkName] as dynamic;
                if (operationInfo == null)
                {
                    return;
                }

                string operationType = operationInfo.OperationType;
                Word.Range operationRange = operationInfo.OperationRange;
                Word.Range originalRange = operationInfo.OriginalRange;
                Word.Range newRange = operationInfo.NewRange;
                
                // 获取句子名称列表（如果存在）
                List<string> originalSentenceNames = null;
                List<string> newSentenceNames = null;
                try
                {
                    originalSentenceNames = operationInfo.OriginalSentenceNames;
                    newSentenceNames = operationInfo.NewSentenceNames;
                }
                catch
                {
                    // 如果不存在，保持为null
                }

                if (action == "accept")
                {
                    ConfirmOperation(operationType, operationRange, originalRange, newRange, originalSentenceNames, newSentenceNames);
                }
                else if (action == "reject")
                {
                    RejectOperation(operationType, operationRange, originalRange, newRange, originalSentenceNames, newSentenceNames);
                }

                // 操作完成后，移除按钮和相关信息
                RemoveInlineButtons(bookmarkName);
            }
            catch (Exception)
            {
                // 忽略处理按钮点击时的异常
            }
        }

        /// <summary>
        /// 删除确认操作对应的历史记录
        /// </summary>
        /// <param name="operationType">操作类型</param>
        /// <param name="originalRange">原始文本范围</param>
        /// <param name="newRange">新文本范围</param>
        private void RemoveHistoricalRecordsByOperation(string operationType, Word.Range originalRange, Word.Range newRange)
        {
            try
            {
                switch (operationType.ToLower())
                {
                    case "insert":
                        // 删除最近的历史插入记录（因为插入操作只添加一个插入记录）
                        var insertRanges = DocumentState.HistoricalInsertRanges;
                        if (insertRanges.Count > 0)
                        {
                            var lastInsert = insertRanges[insertRanges.Count - 1];
                            DocumentState.RemoveHistoricalInsertRange(lastInsert.Start, lastInsert.End);
                        }
                        break;

                    case "replace":
                        // 删除最近的历史删除记录和历史插入记录（因为替换操作添加一个删除记录和一个插入记录）
                        var deleteRanges = DocumentState.HistoricalDeleteRanges;
                        var insertRanges2 = DocumentState.HistoricalInsertRanges;

                        if (deleteRanges.Count > 0)
                        {
                            var lastDelete = deleteRanges[deleteRanges.Count - 1];
                            DocumentState.RemoveHistoricalDeleteRange(lastDelete.Start, lastDelete.End);
                        }

                        if (insertRanges2.Count > 0)
                        {
                            var lastInsert = insertRanges2[insertRanges2.Count - 1];
                            DocumentState.RemoveHistoricalInsertRange(lastInsert.Start, lastInsert.End);
                        }
                        break;

                    case "delete":
                        // 删除最近的历史删除记录（因为删除操作只添加一个删除记录）
                        var deleteRanges2 = DocumentState.HistoricalDeleteRanges;
                        if (deleteRanges2.Count > 0)
                        {
                            var lastDelete = deleteRanges2[deleteRanges2.Count - 1];
                            DocumentState.RemoveHistoricalDeleteRange(lastDelete.Start, lastDelete.End);
                        }
                        break;
                }
            }
            catch (Exception)
            {
                // 忽略删除历史记录时的异常
            }
        }

        /// <summary>
        /// 检查并移除文档中所有新增的按钮（包括接受和丢弃按钮）
        /// </summary>
        /// <returns>返回移除的按钮数量</returns>
        public int RemoveAllPendingButtons()
        {
            var bookmarksToRemove = new List<string>();

            try
            {
                Word.Document document = _application.ActiveDocument;

                // 查找所有书签（按钮）
                foreach (Word.Bookmark bookmark in document.Bookmarks)
                {
                    string bookmarkName = bookmark.Name;
                    if (_operationInfos.ContainsKey(bookmarkName))
                    {
                        // 所有在存储中的按钮都是待处理的（未确认的）
                        bookmarksToRemove.Add(bookmarkName);
                    }
                }

                // 移除找到的所有待处理按钮
                foreach (string bookmarkName in bookmarksToRemove)
                {
                    RemoveInlineButtons(bookmarkName);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 移除了待处理的按钮: {bookmarkName}");
                }

                if (bookmarksToRemove.Count > 0)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 共移除了 {bookmarksToRemove.Count} 个待处理的按钮");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 检查移除按钮时出错: {ex.Message}");
                return 0;
            }

            return bookmarksToRemove.Count;
        }

        /// <summary>
        /// 移除内联按钮
        /// </summary>
        /// <param name="bookmarkName">书签名称</param>
        private void RemoveInlineButtons(string bookmarkName)
        {
            try
            {
                Word.Document document = _application.ActiveDocument;

                // 查找并删除书签（按钮）
                if (document.Bookmarks.Exists(bookmarkName))
                {
                    Word.Bookmark bookmark = document.Bookmarks[bookmarkName];
                    bookmark.Range.Text = ""; // 清空按钮文本
                    bookmark.Delete();
                }

                // 从存储中移除操作信息
                _operationInfos.Remove(bookmarkName);
            }
            catch (Exception)
            {
                // 忽略删除按钮时的异常
            }
        }
    }
}
