using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// MCP 工具 UI 信息：Word 端仅执行；name/alias 来自 Backend Registry。
    /// </summary>
    public static class McpToolsInfo
    {
        public class ToolInfo
        {
            public string name { get; set; }
            public string alias { get; set; }
        }

        private static List<ToolInfo> _allTools = new List<ToolInfo>();
        private static bool _registryLoaded = false;

        public static async Task EnsureRegistryAliasesAsync()
        {
            if (_registryLoaded)
            {
                return;
            }
            await LoadRegistryToolAliasesAsync();
            _registryLoaded = true;
        }

        public static List<ToolInfo> GetAllTools()
        {
            return _allTools.ToList();
        }

        public static string GetToolAlias(string toolName)
        {
            if (string.IsNullOrEmpty(toolName))
            {
                return toolName;
            }

            var tool = _allTools.FirstOrDefault(t => t.name == toolName);
            if (tool != null && !string.IsNullOrEmpty(tool.alias))
            {
                return tool.alias;
            }

            return toolName;
        }

        public static async Task InitializeAsync(object wordApplication = null)
        {
            try
            {
                System.Diagnostics.Debug.WriteLine("[McpToolsInfo] 开始初始化工具信息...");

                _allTools.Clear();

                var toolRegistry = new Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>>();
                McpTools.RegisterBuiltInTools(toolRegistry, wordApplication);

                foreach (var name in toolRegistry.Keys.OrderBy(k => k))
                {
                    _allTools.Add(new ToolInfo { name = name, alias = name });
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[McpToolsInfo] 本地 F_ 执行器 {toolRegistry.Count} 个（metadata 由 Backend Registry 维护）");

                await LoadRegistryToolAliasesAsync();

                if (EasyWriteDiagnostics.IsEnabled(DebugCategory.DocumentExtract))
                {
                    PrintAllTools();
                }

                System.Diagnostics.Debug.WriteLine($"[McpToolsInfo] 工具信息初始化完成，共 {_allTools.Count} 个工具");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[McpToolsInfo] 初始化失败: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"[McpToolsInfo] 异常详情: {ex.StackTrace}");
            }
        }

        private static async Task LoadRegistryToolAliasesAsync()
        {
            try
            {
                string baseUrl = ConfigManager.Config.Api.BaseUrl?.TrimEnd('/');
                if (string.IsNullOrEmpty(baseUrl))
                {
                    return;
                }

                string apiUrl = $"{baseUrl}/api/mcp/tool_aliases";
                System.Diagnostics.Debug.WriteLine($"[McpToolsInfo] 请求 Registry alias: {apiUrl}");

                using (var client = new System.Net.Http.HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(15);
                    var body = await client.GetStringAsync(apiUrl);
                    var response = JsonConvert.DeserializeObject<RegistryAliasesResponse>(body);
                    if (response?.success != true || response.tools == null)
                    {
                        System.Diagnostics.Debug.WriteLine("[McpToolsInfo] Registry alias 响应无效");
                        return;
                    }

                    int merged = 0;
                    foreach (var item in response.tools)
                    {
                        if (string.IsNullOrEmpty(item.name))
                        {
                            continue;
                        }

                        var existing = _allTools.FirstOrDefault(t => t.name == item.name);
                        if (existing != null)
                        {
                            if (!string.IsNullOrEmpty(item.alias))
                            {
                                existing.alias = item.alias;
                                merged++;
                            }
                        }
                        else
                        {
                            _allTools.Add(new ToolInfo
                            {
                                name = item.name,
                                alias = item.alias ?? item.name,
                            });
                            merged++;
                        }
                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[McpToolsInfo] Registry alias 合并 {merged} 条，总计 {_allTools.Count} 个工具");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[McpToolsInfo] 加载 Registry alias 失败: {ex.Message}");
            }
        }

        private class RegistryAliasesResponse
        {
            public bool success { get; set; }
            public List<RegistryAliasItem> tools { get; set; }
            public int count { get; set; }
        }

        private class RegistryAliasItem
        {
            public string name { get; set; }
            public string alias { get; set; }
        }

        private static void PrintAllTools()
        {
            System.Diagnostics.Debug.WriteLine("\n========================================");
            System.Diagnostics.Debug.WriteLine("MCP 工具信息汇总（Backend Registry 为 metadata 权威）");
            System.Diagnostics.Debug.WriteLine("========================================");
            System.Diagnostics.Debug.WriteLine($"总计: {_allTools.Count} 个工具\n");

            foreach (var tool in _allTools.OrderBy(t => t.name))
            {
                System.Diagnostics.Debug.WriteLine($"名称: {tool.name}");
                if (!string.IsNullOrEmpty(tool.alias) && tool.alias != tool.name)
                {
                    System.Diagnostics.Debug.WriteLine($"  别名: {tool.alias}");
                }
            }

            System.Diagnostics.Debug.WriteLine("========================================\n");
        }
    }
}
