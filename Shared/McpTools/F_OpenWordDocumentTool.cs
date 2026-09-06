using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    /// <summary>
    /// 打开已有或新建空白 Word 文档，建立操作渠道并设为默认（B3 / I1）。
    /// </summary>
    public static class F_OpenWordDocumentTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_open_word_document"] = async (args) =>
            {
                try
                {
                    AgentRunCancellation.ThrowIfCancelled();

                    string path = args != null && args.ContainsKey("path")
                        ? args["path"]?.ToString()?.Trim()
                        : null;
                    if (string.IsNullOrEmpty(path))
                    {
                        return Fail("必须提供 path 参数");
                    }

                    bool createBlank = ParseBool(args, "create_blank", false);

                    if (!TryResolveOpenPath(path, createBlank, out string fullPath, out string pathError))
                    {
                        return Fail(pathError);
                    }

                    if (!WordApplicationResolver.TryResolve(
                            wordApplication,
                            out Word.Application app,
                            out string resolveError,
                            createIfMissing: true))
                    {
                        return Fail(resolveError);
                    }

                    // Desktop 启动未注入 Word：此处可能刚 new 出 Application，写回宿主供后续工具使用
                    try
                    {
                        DocumentCheckpointService.SetWordApplication(app);
                        HostCallbacks.RaiseWordApplicationResolved(app);
                    }
                    catch (Exception)
                    {
                    }

                    bool fileExists = File.Exists(fullPath);
                    if (createBlank)
                    {
                        if (fileExists)
                        {
                            return Fail("create_blank=true 但路径已存在，拒绝覆盖: " + fullPath);
                        }

                        string dir = Path.GetDirectoryName(fullPath);
                        if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                        {
                            try
                            {
                                Directory.CreateDirectory(dir);
                            }
                            catch (Exception ex)
                            {
                                return Fail("无法创建目录: " + ex.Message);
                            }
                        }
                    }
                    else if (!fileExists)
                    {
                        return Fail("文件不存在（create_blank=false）: " + fullPath);
                    }

                    Word.Document doc;
                    bool created;
                    try
                    {
                        if (createBlank)
                        {
                            doc = app.Documents.Add();
                            doc.SaveAs2(fullPath, FileFormat: Word.WdSaveFormat.wdFormatDocumentDefault);
                            created = true;
                        }
                        else
                        {
                            doc = FindOpenDocument(app, fullPath) ?? app.Documents.Open(fullPath);
                            created = false;
                        }
                    }
                    catch (Exception ex)
                    {
                        return Fail("打开/新建 Word 文档失败: " + ex.Message);
                    }

                    DocumentState.BindAndActivate(doc);

                    bool indexReady = false;
                    string indexError = null;
                    try
                    {
                        AgentRunCancellation.ThrowIfCancelled();
                        var processing = WordDocumentExtractor.ProcessDocument(
                            doc,
                            ProcessDocumentOptions.ForGetDocumentContent("open_word_document"));
                        indexReady = processing != null;
                    }
                    catch (OperationCanceledException)
                    {
                        return Fail("cancelled by user");
                    }
                    catch (Exception ex)
                    {
                        indexError = ex.Message;
                        System.Diagnostics.Debug.WriteLine(
                            "[F_open_word_document] 建索引失败: " + ex.Message);
                    }

                    WordChannel channel = ChannelRegistry.CreateOrGetWord(doc, fullPath);
                    ChannelRegistry.SetDefault(channel.ChannelId);

                    string name = null;
                    try
                    {
                        name = doc.Name;
                    }
                    catch (Exception)
                    {
                        name = Path.GetFileName(fullPath);
                    }

                    var data = new Dictionary<string, object>
                    {
                        ["channel_id"] = ChannelRegistry.ToPublicId(channel.ChannelId),
                        ["kind"] = "word",
                        ["path"] = fullPath,
                        ["name"] = name ?? "",
                        ["created"] = created,
                        ["index_ready"] = indexReady,
                    };
                    if (!string.IsNullOrEmpty(indexError))
                    {
                        data["index_error"] = indexError;
                    }

                    HostCallbacks.RaiseForegroundDance(new ForegroundDanceRequest
                    {
                        Kind = ForegroundDanceKind.Open,
                        ChannelId = channel.ChannelId,
                        TargetHwnd = ForegroundDanceHwnd.TryResolve(channel.ChannelId),
                        Source = "F_open_word_document",
                    });

                    // 满足 async lambda 签名；实际无 await
                    await Task.CompletedTask;

                    return new ToolResult { Success = true, Data = data };
                }
                catch (OperationCanceledException)
                {
                    return Fail("cancelled by user");
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[ERROR] open_word_document: " + ex.Message);
                    return Fail("打开 Word 文档失败: " + ex.Message);
                }
            };
        }

        private static Word.Document FindOpenDocument(Word.Application app, string fullPath)
        {
            if (app == null || string.IsNullOrEmpty(fullPath))
            {
                return null;
            }

            string norm = NormalizePath(fullPath);
            try
            {
                foreach (Word.Document d in app.Documents)
                {
                    try
                    {
                        string existing = WordChannel.TryReadFullName(d);
                        if (!string.IsNullOrEmpty(existing) &&
                            string.Equals(NormalizePath(existing), norm, StringComparison.OrdinalIgnoreCase))
                        {
                            return d;
                        }
                    }
                    catch (Exception)
                    {
                        // 跳过无效 RCW
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }

            return null;
        }

        private static string NormalizePath(string path)
        {
            try
            {
                return Path.GetFullPath(path).TrimEnd('\\', '/');
            }
            catch (Exception)
            {
                return path?.Trim() ?? "";
            }
        }

        /// <summary>
        /// 解析可打开路径。Windows 文件名禁止 ASCII <c>"</c>，常见备案表却用中文弯引号；
        /// 模型常把路径里的 <c>“”</c> 写成 <c>"</c> 导致 GetFullPath 报非法。此处尝试弯引号/全角变体。
        /// </summary>
        private static bool TryResolveOpenPath(
            string rawPath,
            bool createBlank,
            out string fullPath,
            out string error)
        {
            fullPath = null;
            error = null;

            string path = StripWrappingAsciiQuotes(rawPath?.Trim() ?? "");
            if (string.IsNullOrEmpty(path))
            {
                error = "必须提供 path 参数";
                return false;
            }

            var candidates = new List<string>();
            void AddCandidate(string p)
            {
                if (string.IsNullOrEmpty(p))
                {
                    return;
                }

                foreach (string existing in candidates)
                {
                    if (string.Equals(existing, p, StringComparison.Ordinal))
                    {
                        return;
                    }
                }

                candidates.Add(p);
            }

            AddCandidate(path);
            if (path.IndexOf('"') >= 0)
            {
                AddCandidate(ReplaceAsciiQuotesWithCurly(path));
                AddCandidate(path.Replace("\"", "\uFF02")); // 全角 ＂
            }

            Exception lastIllegal = null;
            foreach (string candidate in candidates)
            {
                string resolved;
                try
                {
                    resolved = Path.GetFullPath(candidate);
                }
                catch (Exception ex)
                {
                    lastIllegal = ex;
                    continue;
                }

                bool exists = File.Exists(resolved);
                if (createBlank)
                {
                    if (!exists)
                    {
                        fullPath = resolved;
                        return true;
                    }

                    continue;
                }

                if (exists)
                {
                    fullPath = resolved;
                    if (!string.Equals(candidate, path, StringComparison.Ordinal))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            "[F_open_word_document] path 已将 ASCII 引号规范为文件系统可用形式: " + resolved);
                    }

                    return true;
                }
            }

            if (createBlank)
            {
                // 所有候选都已存在，或全都非法
                foreach (string candidate in candidates)
                {
                    try
                    {
                        string resolved = Path.GetFullPath(candidate);
                        if (File.Exists(resolved))
                        {
                            error = "create_blank=true 但路径已存在，拒绝覆盖: " + resolved;
                            return false;
                        }

                        fullPath = resolved;
                        return true;
                    }
                    catch (Exception)
                    {
                        // try next
                    }
                }

                error = lastIllegal != null
                    ? "路径非法: " + lastIllegal.Message
                      + "（文件名勿用英文引号 \"，请用中文 “” 或去掉引号）"
                    : "无法解析新建路径";
                return false;
            }

            if (lastIllegal != null && candidates.Count == 1)
            {
                error = "路径非法: " + lastIllegal.Message
                    + "（文件名勿用英文引号 \"，请用中文 “” 或从资源管理器复制真实路径）";
                return false;
            }

            error = "文件不存在（create_blank=false）: " + path
                + (path.IndexOf('"') >= 0
                    ? "。已尝试将英文引号替换为中文/全角引号仍未找到，请核对真实文件名"
                    : "");
            return false;
        }

        private static string StripWrappingAsciiQuotes(string path)
        {
            if (string.IsNullOrEmpty(path) || path.Length < 2)
            {
                return path;
            }

            if (path[0] == '"' && path[path.Length - 1] == '"')
            {
                return path.Substring(1, path.Length - 2).Trim();
            }

            return path;
        }

        /// <summary>按出现次序将 ASCII " 交替替换为 “ 与 ”。</summary>
        private static string ReplaceAsciiQuotesWithCurly(string path)
        {
            if (string.IsNullOrEmpty(path) || path.IndexOf('"') < 0)
            {
                return path;
            }

            var chars = path.ToCharArray();
            bool left = true;
            for (int i = 0; i < chars.Length; i++)
            {
                if (chars[i] != '"')
                {
                    continue;
                }

                chars[i] = left ? '\u201C' : '\u201D';
                left = !left;
            }

            return new string(chars);
        }

        private static bool ParseBool(Dictionary<string, object> args, string key, bool defaultValue)
        {
            if (args == null || !args.ContainsKey(key) || args[key] == null)
            {
                return defaultValue;
            }

            object raw = args[key];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw)?.Trim();
            if (string.IsNullOrEmpty(s))
            {
                return defaultValue;
            }

            if (bool.TryParse(s, out bool parsed))
            {
                return parsed;
            }

            if (s == "1" || string.Equals(s, "yes", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (s == "0" || string.Equals(s, "no", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return defaultValue;
        }

        private static ToolResult Fail(string error)
        {
            return new ToolResult { Success = false, Error = error };
        }
    }
}
