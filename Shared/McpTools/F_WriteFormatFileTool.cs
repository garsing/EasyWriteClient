using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Reflection;
using System.Xml.Linq;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// 格式文件写入工具
    /// 专门用于创建Word表格格式的XML配置文件
    /// </summary>
    public static class F_WriteFormatFileTool
    {
        /// <summary>
        /// 注册格式文件写入工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_write_format_file"] = async (args) =>
            {
                await Task.CompletedTask; // 确保异步执行

                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] write_format_file工具开始执行");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 当前程序集路径: {Assembly.GetExecutingAssembly().Location}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] AppDomain基目录: {AppDomain.CurrentDomain.BaseDirectory}");

                    string content = args.ContainsKey("content") ? args["content"]?.ToString() : "";

                    // 调试：检查接收到的content
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content长度: {content.Length}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content前200字符: {content.Substring(0, Math.Min(200, content.Length))}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content包含\\n: {content.Contains("\\n")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content包含\\r: {content.Contains("\\r")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content包含实际换行: {content.Contains("\n")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 原始content包含实际回车: {content.Contains("\r")}");

                    // 注意：换行符转换已在McpClient.cs中处理，这里不需要重复处理
                    if (content.Contains("\\\""))
                    {
                        content = content.Replace("\\\"", "\"");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 已转换转义双引号\\\"为实际双引号");
                    }
                    if (content.Contains("\\'"))
                    {
                        content = content.Replace("\\'", "'");
                        System.Diagnostics.Debug.WriteLine($"[DEBUG] 已转换转义单引号\\'为实际单引号");
                    }

                    // 调试：检查转换后的content
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 转换后content包含实际换行: {content.Contains("\n")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 转换后content包含实际回车: {content.Contains("\r")}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 转换后content前200字符: {content.Substring(0, Math.Min(200, content.Length))}");

                    if (string.IsNullOrEmpty(content))
                    {
                        return new ToolResult { Success = false, Error = "必须提供content参数（XML格式的表格或图表配置内容）" };
                    }

                    if (!UserService.Instance.CheckLoginStatus()
                        || string.IsNullOrWhiteSpace(UserService.Instance.WorkspaceRootEffective))
                    {
                        return new ToolResult { Success = false, Error = "用户未登录或工作区未初始化，无法创建格式文件" };
                    }

                    // 验证XML格式
                    try
                    {
                        XDocument.Parse(content);
                    }
                    catch (Exception xmlEx)
                    {
                        return new ToolResult { Success = false, Error = $"XML格式无效: {xmlEx.Message}" };
                    }

                    // 生成带时间戳的文件名
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff");
                    string fileName = $"format_{timestamp}.xml";
                    string filePath = WorkspacePathResolver.ResolveWritePath(fileName);

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 写入格式文件: {filePath}");

                    // 写入文件（使用UTF-8编码）
                    File.WriteAllText(filePath, content, Encoding.UTF8);

                    // 上传至云端工作区
                    bool uploadSuccess = await McpToolsHelpers.UploadWorkspaceFileAsync(filePath, fileName);
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 格式文件上传结果: {uploadSuccess}");

                    // 获取文件信息
                    FileInfo fileInfo = new FileInfo(filePath);

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            filename = fileName,  // 返回文件名给大模型
                            file_path = filePath,
                            content_length = content.Length,
                            encoding = "utf-8",
                            file_size = fileInfo.Length,
                            created = fileInfo.CreationTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            modified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            upload_success = uploadSuccess,
                            message = $"格式文件创建成功：{fileName}，共 {content.Length} 个字符，文件大小 {fileInfo.Length} 字节{(uploadSuccess ? " 并同步到云端" : "（同步失败）")}"
                        }
                    };

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] write_format_file工具执行成功，已创建格式文件 {fileName}");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] write_format_file工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"格式文件创建失败: {ex.Message}" };
                }
            };
        }
    }
}
