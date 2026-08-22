using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 文件读取工具
    /// 读取当前会话工作区目录中的文件
    /// </summary>
    public static class F_ReadFileTool
    {
        /// <summary>
        /// 注册文件读取工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_read_file"] = async (args) =>
            {
                await Task.CompletedTask;

                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] read_file工具开始执行");

                    string encoding = args.ContainsKey("encoding") ? args["encoding"]?.ToString().ToLower() : "auto";
                    int maxLength = args.ContainsKey("max_length") ? Convert.ToInt32(args["max_length"]) : -1;
                    int offset = args.ContainsKey("offset") ? Convert.ToInt32(args["offset"]) : 0;

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "filename"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    var read = await FilePathResolver.ReadBytesAsync(resolved).ConfigureAwait(false);
                    if (!read.Success)
                    {
                        return new ToolResult { Success = false, Error = read.Error };
                    }

                    string filename = resolved.Display;
                    string absolutePath = resolved.LocalPath;
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 读取文件: {absolutePath}");

                    FileInfo fileInfo = new FileInfo(absolutePath);
                    string detectedEncoding = encoding;

                    string content;
                    if (encoding == "auto")
                    {
                        using (var reader = new StreamReader(absolutePath, true))
                        {
                            content = reader.ReadToEnd();
                            detectedEncoding = reader.CurrentEncoding.WebName.ToLower();
                        }
                    }
                    else
                    {
                        Encoding fileEncoding;
                        switch (encoding)
                        {
                            case "gb2312":
                                fileEncoding = Encoding.GetEncoding("GB2312");
                                break;
                            case "ascii":
                                fileEncoding = Encoding.ASCII;
                                break;
                            case "unicode":
                                fileEncoding = Encoding.Unicode;
                                break;
                            case "utf-8":
                            default:
                                fileEncoding = Encoding.UTF8;
                                break;
                        }

                        content = File.ReadAllText(absolutePath, fileEncoding);
                        detectedEncoding = encoding;
                    }

                    int actualOffset = Math.Max(0, Math.Min(offset, content.Length));
                    int actualMaxLength = maxLength < 0 ? content.Length - actualOffset : Math.Min(maxLength, content.Length - actualOffset);

                    string resultContent = content.Substring(actualOffset, actualMaxLength);
                    bool isPartial = actualOffset > 0 || actualMaxLength < content.Length;

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            filename = filename,
                            file_path = absolutePath,
                            encoding = detectedEncoding,
                            file_size = fileInfo.Length,
                            total_characters = content.Length,
                            read_characters = resultContent.Length,
                            offset = actualOffset,
                            max_length = maxLength,
                            is_partial = isPartial,
                            content = resultContent,
                            created = fileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            modified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            message = isPartial ?
                                $"成功读取 {filename} 部分内容（{actualOffset}-{actualOffset + actualMaxLength}），共 {resultContent.Length} 个字符" :
                                $"成功读取 {filename} 全部内容，共 {resultContent.Length} 个字符"
                        }
                    };

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] read_file工具执行成功，读取文件 {filename}，共 {resultContent.Length} 个字符");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] read_file工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"文件读取失败: {ex.Message}" };
                }
            };
        }
    }
}
