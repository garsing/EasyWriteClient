using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    /// <summary>
    /// 文件读取：文本回 content；一期图片回 image_base64 走视觉轮。
    /// </summary>
    public static class F_ReadFileTool
    {
        public const int TextWindowChars = 8000;

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

                    string encoding = args != null && args.ContainsKey("encoding")
                        ? args["encoding"]?.ToString().ToLower()
                        : "auto";
                    int maxLength = args != null && args.ContainsKey("max_length") && args["max_length"] != null
                        ? Convert.ToInt32(args["max_length"])
                        : -1;
                    int offset = args != null && args.ContainsKey("offset") && args["offset"] != null
                        ? Convert.ToInt32(args["offset"])
                        : 0;

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "filename"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    var read = await FilePathResolver.ReadBytesAsync(resolved).ConfigureAwait(false);
                    if (!read.Success)
                    {
                        return new ToolResult { Success = false, Error = read.Error };
                    }

                    bool detailed = ReadDetailedFlag(args);
                    FileReadKind kind = FileReadClassifier.Classify(resolved.LocalPath, read.Bytes);
                    if (kind == FileReadKind.RejectedBinary)
                    {
                        return new ToolResult
                        {
                            Success = false,
                            Error = "不能把 " + resolved.Display
                                + " 当文本读。打开稿/表/片请用 F_open_document，看窗口请用 F_capture_document_page_image。"
                        };
                    }

                    if (kind == FileReadKind.Image)
                    {
                        if (detailed)
                        {
                            return new ToolResult { Success = false, Error = "读图片不要传 detailed" };
                        }

                        return ReadImage(args, resolved, read.Bytes);
                    }

                    if (kind == FileReadKind.OfficeDocument)
                    {
                        if (detailed)
                        {
                            return await F_ReadFileKbIngest.RunAsync(resolved.LocalPath).ConfigureAwait(false);
                        }

                        return ReadOfficePlain(resolved, offset, maxLength);
                    }

                    if (detailed)
                    {
                        return new ToolResult { Success = false, Error = "文本文件不要传 detailed" };
                    }

                    return ReadText(resolved, encoding ?? "auto", offset, maxLength);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] read_file工具执行失败: " + ex.Message);
                    return new ToolResult { Success = false, Error = "文件读取失败: " + ex.Message };
                }
            };
        }

        private static ToolResult ReadImage(
            Dictionary<string, object> args,
            ResolvedFilePath resolved,
            byte[] bytes)
        {
            if (HasNonDefaultImageReadArgs(args, out string argError))
            {
                return new ToolResult { Success = false, Error = argError };
            }

            if (bytes == null || bytes.Length == 0)
            {
                return new ToolResult { Success = false, Error = "无法解码图片: " + resolved.Display };
            }

            try
            {
                var image = ImageCaptureCompressor.CompressBytes(bytes);
                return new ToolResult
                {
                    Success = true,
                    Data = new Dictionary<string, object>
                    {
                        ["source_kind"] = "file",
                        ["path"] = resolved.Display,
                        ["format"] = image.Format,
                        ["width_px"] = image.Width,
                        ["height_px"] = image.Height,
                        ["file_size"] = bytes.Length,
                        ["image_base64"] = Convert.ToBase64String(image.Bytes)
                    }
                };
            }
            catch (Exception ex)
            {
                string ext = FileReadClassifier.ExtensionOf(resolved.LocalPath);
                if (string.Equals(ext, "webp", StringComparison.OrdinalIgnoreCase))
                {
                    return new ToolResult
                    {
                        Success = false,
                        Error = "本机无法解码该 webp: " + resolved.Display
                    };
                }

                System.Diagnostics.Debug.WriteLine("[F_read_file] decode image failed: " + ex.Message);
                return new ToolResult { Success = false, Error = "无法解码图片: " + resolved.Display };
            }
        }

        private static bool HasNonDefaultImageReadArgs(Dictionary<string, object> args, out string error)
        {
            error = "读图片不要传 encoding / offset / max_length";
            if (args == null)
            {
                return false;
            }

            if (args.ContainsKey("encoding"))
            {
                string enc = args["encoding"]?.ToString()?.Trim();
                if (!string.IsNullOrEmpty(enc)
                    && !string.Equals(enc, "auto", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            if (args.ContainsKey("offset") && args["offset"] != null
                && Convert.ToInt32(args["offset"]) != 0)
            {
                return true;
            }

            if (args.ContainsKey("max_length") && args["max_length"] != null
                && Convert.ToInt32(args["max_length"]) >= 0)
            {
                return true;
            }

            error = null;
            return false;
        }

        private static bool ReadDetailedFlag(Dictionary<string, object> args)
        {
            if (args == null || !args.ContainsKey("detailed") || args["detailed"] == null)
            {
                return false;
            }

            object raw = args["detailed"];
            if (raw is bool b)
            {
                return b;
            }

            string s = Convert.ToString(raw);
            if (string.IsNullOrWhiteSpace(s))
            {
                return false;
            }

            return string.Equals(s, "true", StringComparison.OrdinalIgnoreCase)
                || s == "1";
        }

        private static ToolResult ReadOfficePlain(ResolvedFilePath resolved, int offset, int maxLength)
        {
            string text;
            try
            {
                text = OfficePlainTextReader.Extract(resolved.LocalPath) ?? "";
            }
            catch (Exception ex)
            {
                return new ToolResult { Success = false, Error = ex.Message };
            }

            if (string.IsNullOrEmpty(text))
            {
                string ext = FileReadClassifier.ExtensionOf(resolved.LocalPath);
                string emptyMsg = string.Equals(ext, "pdf", StringComparison.OrdinalIgnoreCase)
                    ? "没有可抽取的文字，扫描件请 detailed=true"
                    : "没有可抽取的文字";
                return new ToolResult
                {
                    Success = true,
                    Data = new
                    {
                        filename = resolved.Display,
                        file_path = resolved.LocalPath,
                        total_characters = 0,
                        read_characters = 0,
                        offset = 0,
                        range_start = 0,
                        range_end = 0,
                        has_more = false,
                        content = "",
                        message = emptyMsg
                    }
                };
            }

            return SliceTextWindow(resolved, text, offset, maxLength, officePlain: true);
        }

        private static ToolResult ReadText(
            ResolvedFilePath resolved,
            string encoding,
            int offset,
            int maxLength)
        {
            string absolutePath = resolved.LocalPath;
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

            return SliceTextWindow(resolved, content, offset, maxLength, officePlain: false, detectedEncoding);
        }

        private static ToolResult SliceTextWindow(
            ResolvedFilePath resolved,
            string content,
            int offset,
            int maxLength,
            bool officePlain,
            string detectedEncoding = null)
        {
            string filename = resolved.Display;
            string absolutePath = resolved.LocalPath;
            FileInfo fileInfo = new FileInfo(absolutePath);
            content = content ?? "";
            int actualOffset = Math.Max(0, Math.Min(offset, content.Length));
            int window = maxLength <= 0 || maxLength > TextWindowChars
                ? TextWindowChars
                : maxLength;
            int actualMaxLength = Math.Min(window, content.Length - actualOffset);
            string resultContent = content.Substring(actualOffset, actualMaxLength);
            int rangeEnd = actualOffset + actualMaxLength;
            bool hasMore = rangeEnd < content.Length;
            bool isPartial = actualOffset > 0 || hasMore;
            string message;
            if (officePlain && hasMore)
            {
                message = $"只读了 8000 字（第 {actualOffset}-{rangeEnd} 字，全文 {content.Length}）。想读后续请把 offset 接到 {rangeEnd}。";
            }
            else if (isPartial)
            {
                message = $"成功读取 {filename} 第 {actualOffset}-{rangeEnd} 字，全文 {content.Length}；"
                    + (hasMore ? $"还有后面请把 offset 接到 {rangeEnd}" : "本窗已到文末");
            }
            else
            {
                message = $"成功读取 {filename} 全部内容，共 {resultContent.Length} 个字符";
            }

            var data = new Dictionary<string, object>
            {
                ["filename"] = filename,
                ["file_path"] = absolutePath,
                ["file_size"] = fileInfo.Length,
                ["total_characters"] = content.Length,
                ["read_characters"] = resultContent.Length,
                ["offset"] = actualOffset,
                ["range_start"] = actualOffset,
                ["range_end"] = rangeEnd,
                ["has_more"] = hasMore,
                ["max_length"] = window,
                ["is_partial"] = isPartial,
                ["content"] = resultContent,
                ["created"] = fileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"),
                ["modified"] = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                ["message"] = message
            };
            if (!officePlain && detectedEncoding != null)
            {
                data["encoding"] = detectedEncoding;
            }

            return new ToolResult { Success = true, Data = data };
        }
    }
}
