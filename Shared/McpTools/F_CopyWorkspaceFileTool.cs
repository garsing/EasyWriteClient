using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace WordAddIn1
{
    public static class F_CopyWorkspaceFileTool
    {
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_copy_workspace_file"] = args =>
            {
                try
                {
                    string sourceRaw = FilePathResolver.TryGetArg(args, "source");
                    string destRaw = FilePathResolver.TryGetArg(args, "dest");
                    if (string.IsNullOrEmpty(sourceRaw))
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "必须提供 source 参数" });
                    }

                    if (string.IsNullOrEmpty(destRaw))
                    {
                        return Task.FromResult(new ToolResult { Success = false, Error = "必须提供 dest 参数" });
                    }

                    if (destRaw.EndsWith("/") || destRaw.EndsWith("\\"))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "目标必须是文件路径，不能是目录"
                        });
                    }

                    if (!FilePathResolver.TryResolve(sourceRaw, out ResolvedFilePath source, out string sourceError))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = RelabelPathError(sourceError, "source")
                        });
                    }

                    if (!FilePathResolver.TryResolve(destRaw, out ResolvedFilePath dest, out string destError))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = RelabelPathError(destError, "dest")
                        });
                    }

                    if (source.Kind == dest.Kind)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = source.Kind == FilePathKind.Absolute
                                ? "两端都是绝对路径；本工具只做区外绝对路径与工作区相对名之间的复制"
                                : "两端都是工作区路径；区内互拷请用 F_read_file / F_write_file"
                        });
                    }

                    if (source.Kind == FilePathKind.Workspace && !ConversationContext.WorkspaceReady)
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "会话工作区未对齐"
                        });
                    }

                    if (Directory.Exists(source.LocalPath) && !File.Exists(source.LocalPath))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "只拷文件，不拷目录: " + source.Display
                        });
                    }

                    if (!File.Exists(source.LocalPath))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "源文件不存在: " + source.Display
                        });
                    }

                    if (Directory.Exists(dest.LocalPath))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "目标已存在且是目录: " + dest.Display
                        });
                    }

                    if (string.Equals(
                        Path.GetFullPath(source.LocalPath),
                        Path.GetFullPath(dest.LocalPath),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return Task.FromResult(new ToolResult
                        {
                            Success = false,
                            Error = "源与目标是同一文件"
                        });
                    }

                    string parent = Path.GetDirectoryName(dest.LocalPath);
                    if (!string.IsNullOrEmpty(parent))
                    {
                        Directory.CreateDirectory(parent);
                    }

                    File.Copy(source.LocalPath, dest.LocalPath, overwrite: true);

                    var info = new FileInfo(dest.LocalPath);
                    string direction = dest.Kind == FilePathKind.Workspace ? "in" : "out";
                    return Task.FromResult(new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            source = source.Display,
                            dest = dest.Display,
                            direction,
                            bytes = info.Exists ? info.Length : 0
                        }
                    });
                }
                catch (Exception ex)
                {
                    return Task.FromResult(new ToolResult
                    {
                        Success = false,
                        Error = "复制文件失败: " + ex.Message
                    });
                }
            };
        }

        private static string RelabelPathError(string error, string param)
        {
            if (string.IsNullOrEmpty(error))
            {
                return "路径非法: " + param;
            }

            if (error.IndexOf("必须提供 path", StringComparison.Ordinal) >= 0)
            {
                return "必须提供 " + param + " 参数";
            }

            return error;
        }
    }
}
