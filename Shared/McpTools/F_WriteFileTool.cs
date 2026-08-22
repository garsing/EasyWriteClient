using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace WordAddIn1
{
    public static class F_WriteFileTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_write_file"] = async (args) =>
            {
                try
                {
                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    string content = args != null && args.ContainsKey("content")
                        ? args["content"]?.ToString()
                        : null;
                    if (content == null)
                    {
                        return new ToolResult { Success = false, Error = "必须提供 content 参数" };
                    }

                    if (string.Equals(Path.GetExtension(resolved.LocalPath), ".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            XDocument.Parse(content);
                        }
                        catch (Exception xmlEx)
                        {
                            return new ToolResult { Success = false, Error = $"XML格式无效: {xmlEx.Message}" };
                        }
                    }

                    var written = await FilePathResolver.WriteAsync(resolved, content, Encoding.UTF8).ConfigureAwait(false);
                    if (!written.Success)
                    {
                        return new ToolResult { Success = false, Error = written.Error };
                    }

                    var info = new FileInfo(resolved.LocalPath);
                    return new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            path = resolved.Display,
                            kind = resolved.Kind == FilePathKind.Workspace ? "workspace" : "absolute",
                            bytes = info.Exists ? info.Length : 0,
                            encoding = "utf-8",
                            message = resolved.Kind == FilePathKind.Workspace
                                ? $"已写入工作区并同步: {resolved.Display}"
                                : $"已写入本机: {resolved.Display}"
                        }
                    };
                }
                catch (Exception ex)
                {
                    return new ToolResult { Success = false, Error = $"写入文件失败: {ex.Message}" };
                }
            };
        }
    }
}
