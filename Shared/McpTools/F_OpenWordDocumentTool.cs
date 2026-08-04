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

                    string fullPath;
                    try
                    {
                        fullPath = Path.GetFullPath(path);
                    }
                    catch (Exception ex)
                    {
                        return Fail("路径非法: " + ex.Message);
                    }

                    if (!WordApplicationResolver.TryResolve(
                            wordApplication,
                            out Word.Application app,
                            out string resolveError,
                            createIfMissing: true))
                    {
                        return Fail(resolveError);
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

                    try
                    {
                        doc.Activate();
                    }
                    catch (Exception)
                    {
                        // 激活失败不阻断建渠道
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
                        ["channel_id"] = channel.ChannelId,
                        ["kind"] = "word",
                        ["doc_uuid"] = channel.DocUuid,
                        ["path"] = fullPath,
                        ["name"] = name ?? "",
                        ["created"] = created,
                        ["index_ready"] = indexReady,
                    };
                    if (!string.IsNullOrEmpty(indexError))
                    {
                        data["index_error"] = indexError;
                    }

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
