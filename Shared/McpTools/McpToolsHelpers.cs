using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// MCP工具辅助方法
    /// 提供各种工具使用的通用辅助方法
    /// </summary>
    public static class McpToolsHelpers
    {
        /// <summary>
        /// 在指定范围内搜索文本
        /// </summary>
        /// <param name="searchRange">搜索范围</param>
        /// <param name="searchText">要搜索的文本</param>
        /// <returns>找到的范围，如果没找到则返回null</returns>
        public static Word.Range FindText(Word.Range searchRange, string searchText)
        {
            System.Diagnostics.Debug.WriteLine($"[DEBUG] FindText called with searchText: '{searchText}'");

            if (string.IsNullOrEmpty(searchText))
            {
                System.Diagnostics.Debug.WriteLine("[DEBUG] Search text is null or empty");
                return null;
            }

            try
            {
                Word.Find find = searchRange.Find;
                find.ClearFormatting(); // 清除之前的格式设置
                find.Text = searchText;
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.MatchCase = false; // 不区分大小写
                find.MatchWholeWord = false; // 不要求全词匹配
                find.MatchWildcards = false; // 不使用通配符
                find.MatchSoundsLike = false; // 不匹配发音相似的单词
                find.MatchAllWordForms = false; // 不匹配所有词形

                System.Diagnostics.Debug.WriteLine($"[DEBUG] Executing find for text: '{searchText}'");
                bool found = find.Execute();

                if (found)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] Text '{searchText}' found successfully at position {searchRange.Start}");
                    return searchRange.Duplicate;
                }

                System.Diagnostics.Debug.WriteLine($"[DEBUG] Text '{searchText}' not found");
                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DEBUG] FindText error: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 获取指定位置之前的上下文内容
        /// </summary>
        public static string GetContextBefore(Word.Document doc, int position, int maxLength)
        {
            try
            {
                int start = Math.Max(0, position - maxLength);
                Word.Range range = doc.Range(start, position);
                string text = range.Text?.Trim() ?? "";
                return text.Length > maxLength ? text.Substring(text.Length - maxLength) : text;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 获取指定位置之后的上下文内容
        /// </summary>
        public static string GetContextAfter(Word.Document doc, int position, int maxLength)
        {
            try
            {
                int end = Math.Min(doc.Content.End, position + maxLength);
                Word.Range range = doc.Range(position, end);
                string text = range.Text?.Trim() ?? "";
                return text.Length > maxLength ? text.Substring(0, maxLength) : text;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 提取表格标题
        /// </summary>
        public static string ExtractTableTitle(Word.Document doc, int tableStart, int maxLength)
        {
            try
            {
                // 向前查找可能的标题
                int searchStart = Math.Max(0, tableStart - 200); // 向前搜索200个字符
                Word.Range titleRange = doc.Range(searchStart, tableStart);

                // 查找包含"表"、"Table"等关键词的段落
                string titleText = titleRange.Text ?? "";
                var lines = titleText.Split('\n').Reverse(); // 从后往前查找

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (trimmed.Contains("表") || trimmed.Contains("Table") ||
                        trimmed.Contains("表格") || trimmed.Contains("数据"))
                    {
                        return trimmed.Length > maxLength ? trimmed.Substring(0, maxLength) : trimmed;
                    }
                }
            }
            catch { }

            return "";
        }

        /// <summary>
        /// 提取图片说明文字
        /// </summary>
        public static string ExtractImageCaption(Word.Document doc, int imageStart, int maxLength)
        {
            try
            {
                // 向后查找图片说明
                int searchEnd = Math.Min(doc.Content.End, imageStart + 200); // 向后搜索200个字符
                Word.Range captionRange = doc.Range(imageStart, searchEnd);

                string captionText = captionRange.Text ?? "";
                var lines = captionText.Split('\n');

                foreach (var line in lines)
                {
                    string trimmed = line.Trim();
                    if (trimmed.Contains("图") || trimmed.Contains("Figure") ||
                        trimmed.Contains("图片") || trimmed.Contains("插图"))
                    {
                        return trimmed.Length > maxLength ? trimmed.Substring(0, maxLength) : trimmed;
                    }
                }
            }
            catch { }

            return "";
        }

        /// <summary>
        /// 清理JSON字符串，确保它是有效的JSON格式
        /// </summary>
        /// <param name="jsonString">需要清理的JSON字符串</param>
        /// <returns>清理后的JSON字符串</returns>
        public static string CleanJsonString(string jsonString)
        {
            if (string.IsNullOrEmpty(jsonString))
                return jsonString;

            // 查找第一个完整的JSON对象（从{开始到}结束）
            int startIndex = jsonString.IndexOf('{');
            if (startIndex == -1)
                return jsonString;

            int braceCount = 0;
            int endIndex = -1;

            for (int i = startIndex; i < jsonString.Length; i++)
            {
                if (jsonString[i] == '{')
                {
                    braceCount++;
                }
                else if (jsonString[i] == '}')
                {
                    braceCount--;
                    if (braceCount == 0)
                    {
                        endIndex = i;
                        break;
                    }
                }
            }

            if (endIndex != -1 && endIndex >= startIndex)
            {
                string cleanJson = jsonString.Substring(startIndex, endIndex - startIndex + 1);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 清理后的JSON: {cleanJson}");
                return cleanJson;
            }

            return jsonString;
        }

        /// <summary>
        /// 获取用户工作目录路径
        /// 根据运行环境（开发/生产）自动确定项目根目录，然后返回用户特定的工作目录
        /// </summary>
        /// <param name="subDirectory">子目录名称（如 "format_files", "curr_docs" 等）</param>
        /// <param name="username">用户名，如果为空则从UserService获取</param>
        /// <returns>用户工作目录的完整路径</returns>
        public static string GetUserWorkDirectory(string subDirectory, string username = null)
        {
            // 获取用户名
            if (string.IsNullOrEmpty(username))
            {
                var userService = UserService.Instance;
                username = userService.IsLoggedIn ? userService.UserName : "anonymous";
            }

            // 如果用户名为空或anonymous，返回null（表示无法确定工作目录）
            if (string.IsNullOrEmpty(username) || username == "anonymous")
            {
                return null;
            }

            // 构建工作目录路径（相对路径）
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string workDirectoryPath;

            // 在开发环境中，AppDomain.CurrentDomain.BaseDirectory 指向 bin\Debug
            // 我们需要向上两级目录到达项目根目录
            if (baseDir.Contains("bin\\Debug") || baseDir.Contains("bin/Debug"))
            {
                // 从 bin\Debug 向上两级到达项目根目录
                string debugDir = Path.GetDirectoryName(baseDir.TrimEnd('\\', '/'));
                string projectRoot = Path.GetDirectoryName(debugDir);
                workDirectoryPath = Path.Combine(projectRoot, "tmp", username, subDirectory);
            }
            else
            {
                // 在生产环境中，使用标准的相对路径
                workDirectoryPath = Path.Combine(baseDir, "tmp", username, subDirectory);
            }

            // 确保目录存在
            if (!Directory.Exists(workDirectoryPath))
            {
                Directory.CreateDirectory(workDirectoryPath);
                System.Diagnostics.Debug.WriteLine($"[DEBUG] 创建用户工作目录: {workDirectoryPath}");
            }

            return workDirectoryPath;
        }

        /// <summary>
        /// 获取项目根目录路径
        /// </summary>
        /// <returns>项目根目录的完整路径</returns>
        public static string GetProjectRootDirectory()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;

            // 在开发环境中，AppDomain.CurrentDomain.BaseDirectory 指向 bin\Debug
            // 我们需要向上两级目录到达项目根目录
            if (baseDir.Contains("bin\\Debug") || baseDir.Contains("bin/Debug"))
            {
                // 从 bin\Debug 向上两级到达项目根目录
                string debugDir = Path.GetDirectoryName(baseDir.TrimEnd('\\', '/'));
                return Path.GetDirectoryName(debugDir);
            }
            else
            {
                // 在生产环境中，BaseDirectory就是项目根目录
                return baseDir;
            }
        }

        /// <summary>
        /// 工作区 .yaml/.yml/.xml 在资源管理器中默认隐藏（不影响程序读写）。
        /// </summary>
        public static void TryHideWorkspaceFormatFile(string localPath)
        {
            if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
            {
                return;
            }

            string ext = Path.GetExtension(localPath);
            if (!string.Equals(ext, ".yaml", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(ext, ".yml", StringComparison.OrdinalIgnoreCase)
                && !string.Equals(ext, ".xml", StringComparison.OrdinalIgnoreCase))
            {
                return;
            }

            try
            {
                var attrs = File.GetAttributes(localPath);
                if ((attrs & FileAttributes.Hidden) == 0)
                {
                    File.SetAttributes(localPath, attrs | FileAttributes.Hidden);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[McpToolsHelpers] 设置格式文件隐藏属性失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从用户工作区读取文件（本地优先；云端同名且更新则覆盖后再读）。
        /// </summary>
        public static async Task<(bool success, string content, string error)> LoadUserFormatFileAsync(string filename)
        {
            var ensure = await EnsureWorkspaceFileAsync(filename).ConfigureAwait(false);
            if (!ensure.success)
            {
                return (false, null, ensure.error);
            }

            try
            {
                string content;
                using (var reader = new StreamReader(ensure.localPath, true))
                {
                    content = reader.ReadToEnd();
                }

                if (string.IsNullOrEmpty(content))
                {
                    return (false, null, $"文件内容为空: {filename}");
                }

                return (true, content, null);
            }
            catch (Exception ex)
            {
                return (false, null, $"读取文件失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 确保工作区文件存在于本地（ResolveReadPath → 若云端同名且更新则覆盖下载 → 否则本地缺失时下载）。
        /// </summary>
        public static async Task<(bool success, string localPath, string error)> EnsureWorkspaceFileAsync(
            string filename,
            string conversationId = null)
        {
            // 后端略新于本地时才覆盖；吸收上传往返与文件系统时间精度抖动
            const double NewerSkewSeconds = 1.0;

            try
            {
                if (string.IsNullOrWhiteSpace(filename))
                {
                    return (false, null, "文件名不能为空");
                }

                if (!UserService.Instance.CheckLoginStatus())
                {
                    return (false, null, "用户未登录，无法访问工作区文件");
                }

                if (string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
                {
                    return (false, null, "工作区未初始化，请先在设置页配置工作目录");
                }

                string relativePath = WorkspacePathResolver.BuildRelativePath(
                    conversationId ?? ConversationContext.CurrentId,
                    filename);

                string localPath = WorkspacePathResolver.ResolveReadPath(filename, conversationId);
                bool localExists = !string.IsNullOrEmpty(localPath) && File.Exists(localPath);

                if (localExists)
                {
                    try
                    {
                        var (remoteExists, remoteStat) = await BackendApiClient
                            .TryGetWorkspaceFileStatAsync(relativePath)
                            .ConfigureAwait(false);

                        if (remoteExists && remoteStat != null)
                        {
                            DateTime localUtc = File.GetLastWriteTimeUtc(localPath);
                            if (remoteStat.MtimeUtc > localUtc.AddSeconds(NewerSkewSeconds))
                            {
                                System.Diagnostics.Debug.WriteLine(
                                    $"[McpToolsHelpers] 云端更新，覆盖本地: {filename} remote={remoteStat.MtimeUtc:o} local={localUtc:o}");

                                bool refreshed = await FileDownloader
                                    .DownloadFileFromUserDirectory(relativePath, localPath)
                                    .ConfigureAwait(false);

                                if (refreshed && File.Exists(localPath))
                                {
                                    TryAlignLocalMtime(localPath, remoteStat.MtimeUtc);
                                    TryHideWorkspaceFormatFile(localPath);
                                    return (true, localPath, null);
                                }

                                System.Diagnostics.Debug.WriteLine(
                                    $"[McpToolsHelpers] 云端更新但下载失败，沿用本地: {filename}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[McpToolsHelpers] 比对云端 mtime 失败，沿用本地: {ex.Message}");
                    }

                    TryHideWorkspaceFormatFile(localPath);
                    return (true, localPath, null);
                }

                localPath = WorkspacePathResolver.ResolveWritePath(filename, conversationId);

                bool downloaded = await FileDownloader.DownloadFileFromUserDirectory(relativePath, localPath)
                    .ConfigureAwait(false);

                if (downloaded || File.Exists(localPath))
                {
                    TryHideWorkspaceFormatFile(localPath);
                    return (true, localPath, null);
                }

                return (false, null, $"文件不存在: {filename}");
            }
            catch (Exception ex)
            {
                return (false, null, $"获取工作区文件失败: {ex.Message}");
            }
        }

        private static void TryAlignLocalMtime(string localPath, DateTime remoteMtimeUtc)
        {
            try
            {
                DateTime utc = remoteMtimeUtc.Kind == DateTimeKind.Utc
                    ? remoteMtimeUtc
                    : remoteMtimeUtc.ToUniversalTime();
                File.SetLastWriteTimeUtc(localPath, utc);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[McpToolsHelpers] 对齐本地 mtime 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 上传本地文件至云端工作区 sessions/{id}/。
        /// </summary>
        public static Task<bool> UploadWorkspaceFileAsync(
            string localFilePath,
            string filename,
            string conversationId = null)
        {
            TryHideWorkspaceFormatFile(localFilePath);
            string relativePath = WorkspacePathResolver.BuildRelativePath(
                conversationId ?? ConversationContext.CurrentId,
                filename);
            return BackendApiClient.UploadUserFileAsync(localFilePath, relativePath);
        }

        /// <summary>
        /// 当前会话工作目录（扁平，不分 format_files/csv）。
        /// </summary>
        public static string GetWorkspaceSessionDirectory(string conversationId = null)
        {
            return WorkspacePathResolver.GetSessionDirectory(conversationId);
        }
    }
}
