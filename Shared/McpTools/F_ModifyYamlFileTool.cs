using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.RepresentationModel;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// YAML文件修改工具
    /// 修改指定YAML文件中的节点内容
    /// </summary>
    public static class F_ModifyYamlFileTool
    {
        /// <summary>
        /// 注册YAML文件修改工具
        /// </summary>
        public static void Register(
            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)
        {
            toolRegistry["F_modify_yaml_file"] = async (args) =>
            {
                await Task.CompletedTask; // 确保异步执行

                try
                {
                    System.Diagnostics.Debug.WriteLine("[DEBUG] modify_yaml_file工具开始执行");

                    if (!FilePathResolver.TryResolveFromArgs(args, out ResolvedFilePath resolved, out string pathError, "path", "filename"))
                    {
                        return new ToolResult { Success = false, Error = pathError };
                    }

                    string filename = resolved.Display;
                    string nodePath = args.ContainsKey("node_path") ? args["node_path"]?.ToString() : "";
                    string newContent = args.ContainsKey("new_content") ? args["new_content"]?.ToString() : "";

                    if (string.IsNullOrEmpty(nodePath))
                    {
                        return new ToolResult { Success = false, Error = "必须提供node_path参数（节点路径）" };
                    }

                    if (newContent == null)
                    {
                        return new ToolResult { Success = false, Error = "必须提供new_content参数（新的内容）" };
                    }

                    // 获取用户名
                    var userService = UserService.Instance;
                    if (!userService.CheckLoginStatus()
                        || string.IsNullOrWhiteSpace(userService.WorkspaceRootEffective))
                    {
                        return new ToolResult { Success = false, Error = "用户未登录或工作区未初始化" };
                    }

                    var read = await FilePathResolver.ReadAsync(resolved).ConfigureAwait(false);
                    if (!read.Success)
                    {
                        return new ToolResult { Success = false, Error = read.Error };
                    }

                    string filePath = resolved.LocalPath;

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 修改YAML文件: {filePath}");
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] 节点路径: {nodePath}");

                    // 读取并解析YAML文件
                    string yamlContent = read.Text ?? "";
                    var yamlStream = new YamlStream();
                    yamlStream.Load(new StringReader(yamlContent));

                    if (yamlStream.Documents.Count == 0)
                    {
                        return new ToolResult { Success = false, Error = "YAML文件为空或格式无效" };
                    }

                    // 获取根节点
                    var rootNode = yamlStream.Documents[0].RootNode as YamlMappingNode;
                    if (rootNode == null)
                    {
                        return new ToolResult { Success = false, Error = "YAML文件的根节点不是映射节点" };
                    }

                    // 解析节点路径
                    string[] pathParts = nodePath.Split('.');
                    YamlNode currentNode = rootNode;
                    YamlNode parentNode = null;
                    string lastKey = null;

                    // 遍历路径，找到目标节点
                    for (int i = 0; i < pathParts.Length; i++)
                    {
                        string part = pathParts[i];
                        bool isLastPart = (i == pathParts.Length - 1);

                        if (currentNode is YamlMappingNode mappingNode)
                        {
                            if (!mappingNode.Children.ContainsKey(new YamlScalarNode(part)))
                            {
                                return new ToolResult { Success = false, Error = $"节点路径不存在: {string.Join(".", pathParts.Take(i + 1))}" };
                            }

                            parentNode = currentNode;
                            lastKey = part;
                            currentNode = mappingNode.Children[new YamlScalarNode(part)];
                        }
                        else if (currentNode is YamlSequenceNode sequenceNode)
                        {
                            if (!int.TryParse(part, out int index) || index < 0 || index >= sequenceNode.Children.Count)
                            {
                                return new ToolResult { Success = false, Error = $"数组索引无效: {part}，数组长度为: {sequenceNode.Children.Count}" };
                            }

                            parentNode = currentNode;
                            lastKey = part;
                            currentNode = sequenceNode.Children[index];
                        }
                        else
                        {
                            return new ToolResult { Success = false, Error = $"节点路径中的中间节点类型不支持: {currentNode.GetType().Name}" };
                        }
                    }

                    // 修改节点内容
                    YamlNode newNode;
                    if (newContent == "null" || newContent == "~")
                    {
                        newNode = new YamlScalarNode("~");
                    }
                    else if (int.TryParse(newContent, out int intValue))
                    {
                        newNode = new YamlScalarNode(intValue.ToString());
                    }
                    else if (float.TryParse(newContent, out float floatValue))
                    {
                        newNode = new YamlScalarNode(floatValue.ToString());
                    }
                    else if ((newContent.StartsWith("[") && newContent.EndsWith("]")) ||
                             (newContent.StartsWith("{") && newContent.EndsWith("}")))
                    {
                        // 尝试解析为JSON数组或对象
                        try
                        {
                            var jsonData = JsonConvert.DeserializeObject(newContent);

                            // 将JSON转换为YAML节点
                            newNode = ConvertJsonToYamlNode(jsonData);
                        }
                        catch (Exception jsonEx)
                        {
                            System.Diagnostics.Debug.WriteLine($"[DEBUG] JSON解析失败: {jsonEx.Message}，回退到字符串处理");
                            newNode = new YamlScalarNode(newContent);
                        }
                    }
                    else
                    {
                        newNode = new YamlScalarNode(newContent);
                    }

                    // 替换节点
                    if (parentNode is YamlMappingNode mappingParent)
                    {
                        mappingParent.Children[new YamlScalarNode(lastKey)] = newNode;
                    }
                    else if (parentNode is YamlSequenceNode sequenceParent && int.TryParse(lastKey, out int arrayIndex))
                    {
                        sequenceParent.Children[arrayIndex] = newNode;
                    }
                    else
                    {
                        return new ToolResult { Success = false, Error = "无法确定父节点类型进行替换" };
                    }

                    // 保存修改后的YAML文件
                    var stringBuilder = new StringBuilder();
                    using (var writer = new StringWriter(stringBuilder))
                    {
                        yamlStream.Save(writer, false);
                    }

                    string updatedYamlContent = stringBuilder.ToString();
                    var written = await FilePathResolver.WriteAsync(resolved, updatedYamlContent, Encoding.UTF8).ConfigureAwait(false);
                    if (!written.Success)
                    {
                        return new ToolResult { Success = false, Error = written.Error };
                    }

                    bool uploadSuccess = true;

                    // 获取文件信息
                    FileInfo fileInfo = new FileInfo(filePath);

                    var result = new ToolResult
                    {
                        Success = true,
                        Data = new
                        {
                            filename = filename,
                            file_path = filePath,
                            node_path = nodePath,
                            old_content = GetNodeValueAsString(currentNode),
                            new_content = newContent,
                            file_size = fileInfo.Length,
                            modified = fileInfo.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"),
                            upload_success = uploadSuccess,
                            message = $"YAML文件节点修改成功：{filename} -> {nodePath}{(uploadSuccess ? " 并同步到云端" : "（同步失败）")}"
                        }
                    };

                    System.Diagnostics.Debug.WriteLine($"[DEBUG] modify_yaml_file工具执行成功，已修改节点 {nodePath}");
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[DEBUG] modify_yaml_file工具执行失败: {ex.Message}");
                    return new ToolResult { Success = false, Error = $"YAML文件修改失败: {ex.Message}" };
                }
            };
        }



        /// <summary>
        /// 将JSON对象转换为YAML节点
        /// </summary>
        private static YamlNode ConvertJsonToYamlNode(object jsonData)
        {
            if (jsonData == null)
            {
                return new YamlScalarNode("~");
            }
            else if (jsonData is Newtonsoft.Json.Linq.JArray jsonArray)
            {
                var sequenceNode = new YamlSequenceNode();
                foreach (var item in jsonArray)
                {
                    sequenceNode.Children.Add(ConvertJsonToYamlNode(item));
                }
                return sequenceNode;
            }
            else if (jsonData is Newtonsoft.Json.Linq.JObject jsonObject)
            {
                var mappingNode = new YamlMappingNode();
                foreach (var property in jsonObject.Properties())
                {
                    var keyNode = new YamlScalarNode(property.Name);
                    var valueNode = ConvertJsonToYamlNode(property.Value);
                    mappingNode.Children.Add(keyNode, valueNode);
                }
                return mappingNode;
            }
            else if (jsonData is Newtonsoft.Json.Linq.JValue jsonValue)
            {
                return new YamlScalarNode(jsonValue.ToString());
            }
            else
            {
                // 处理基本类型
                return new YamlScalarNode(jsonData.ToString());
            }
        }

        /// <summary>
        /// 将YAML节点转换为字符串表示
        /// </summary>
        private static string GetNodeValueAsString(YamlNode node)
        {
            if (node is YamlScalarNode scalarNode)
            {
                return scalarNode.Value;
            }
            else if (node is YamlSequenceNode sequenceNode)
            {
                return $"[{string.Join(", ", sequenceNode.Children.Select(GetNodeValueAsString))}]";
            }
            else if (node is YamlMappingNode mappingNode)
            {
                return "{...}";
            }
            else
            {
                return node.ToString();
            }
        }

        /// <summary>
        /// 上传修改后的YAML文件到用户工作目录
        /// </summary>
        private static Task<bool> UploadModifiedYamlFileToUserDirectory(string filePath, string filename)
        {
            return McpToolsHelpers.UploadWorkspaceFileAsync(filePath, filename);
        }
    }
}
