using System;

using System.Collections.Generic;

using System.Threading.Tasks;

using Newtonsoft.Json;

using Newtonsoft.Json.Linq;

using WordAddIn1.DocumentHost;



namespace WordAddIn1

{

    /// <summary>

    /// 获取当前文档指定分块的 display_content 工具

    /// 根据 chunk_idx 从 API 获取当前文档（DocumentState）对应分块的 display_content

    /// 渠道解析经 <see cref="DocumentHostAdapter"/>（Word/WPS）；只读不抢前台。

    /// </summary>

    public static class F_GetCurrDocChunkDisplayContentTool

    {

        /// <summary>

        /// 注册获取当前文档分块 display_content 工具

        /// </summary>

        public static void Register(

            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,
            object wordApplication)

        {
            toolRegistry["F_get_curr_doc_chunk_display_content"] = async (args) =>

            {

                try

                {

                    // 只读：绑定渠道会话即可，勿 Activate 文档窗口
                    if (!DocumentHostAdapter.TryResolveContext(
                            args,
                            wordApplication,
                            out DocumentSessionContext context,
                            out ToolResult resolveError,
                            activateDocument: false))

                    {

                        return resolveError;

                    }

                    System.Diagnostics.Debug.WriteLine(
                        $"[F_get_curr_doc_chunk_display_content] host={context.Kind.ToString().ToLowerInvariant()}, channel_id={context.ChannelId}");



                    // 获取参数

                    if (!args.ContainsKey("chunk_idx"))

                    {

                        return new ToolResult { Success = false, Error = "缺少必需参数: chunk_idx" };

                    }



                    object chunkIdxObj = args["chunk_idx"];

                    int chunkIdx;

                    

                    // 尝试转换chunk_idx为整数

                    if (chunkIdxObj is int)

                    {

                        chunkIdx = (int)chunkIdxObj;

                    }

                    else if (chunkIdxObj is long)

                    {

                        chunkIdx = (int)(long)chunkIdxObj;

                    }

                    else if (int.TryParse(chunkIdxObj?.ToString(), out int parsedChunkIdx))

                    {

                        chunkIdx = parsedChunkIdx;

                    }

                    else

                    {

                        return new ToolResult { Success = false, Error = "chunk_idx必须是整数" };

                    }



                    // 验证chunk_idx必须大于等于1

                    if (chunkIdx < 1)

                    {

                        return new ToolResult { Success = false, Error = "chunk_idx必须大于等于1" };

                    }



                    // 从DocumentState读取doc_uuid和SnapshotIdx

                    string docUuid = DocumentState.CurrDocUuid;

                    int snapshotIdx = DocumentState.SnapshotIdx;



                    // 验证doc_uuid是否存在

                    if (string.IsNullOrEmpty(docUuid))

                    {

                        return new ToolResult { Success = false, Error = "未找到文档UUID，请先处理文档" };

                    }



                    // 验证snapshot_idx是否有效（必须大于等于1）

                    if (snapshotIdx < 1)

                    {

                        return new ToolResult { Success = false, Error = "快照索引无效，请先处理文档" };

                    }



                    // 获取用户服务

                    var userService = UserService.Instance;

                    if (!userService.CheckLoginStatus())

                    {

                        return new ToolResult { Success = false, Error = "用户未登录，无法调用接口" };

                    }



                    // 构建API URL

                    string baseUrl = ConfigManager.Config.Api.BaseUrl;

                    string apiUrl = $"{baseUrl}/curr_doc/display_content?doc_uuid={Uri.EscapeDataString(docUuid)}&snapshot_idx={snapshotIdx}&chunk_idx={chunkIdx}";



                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] 调用API: {apiUrl}");

                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] doc_uuid: {docUuid}, snapshot_idx: {snapshotIdx}, chunk_idx: {chunkIdx}");



                    var httpResult = await BackendApiClient.GetAuthenticatedAsync(

                        apiUrl,

                        retryOnUnauthorized: true,

                        TimeSpan.FromSeconds(30));

                    var responseContent = httpResult.Body;



                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] 响应状态码: {httpResult.StatusCode}");

                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] 响应内容: {responseContent}");



                    if (httpResult.Success)

                    {

                            // 解析响应

                            var responseObj = JObject.Parse(responseContent);

                            if (responseObj["success"]?.Value<bool>() == true && responseObj["data"] != null)

                            {

                                var data = responseObj["data"];

                                string displayContent = data["display_content"]?.Value<string>() ?? "";



                                System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] ✅ 成功获取display_content，长度: {displayContent.Length} 字符");



                                // 构建返回数据

                                var resultData = new Dictionary<string, object>

                                {

                                    { "display_content", displayContent },

                                    { "chunk_idx", chunkIdx },

                                    { "doc_uuid", docUuid },

                                    { "snapshot_idx", snapshotIdx }

                                };



                                return new ToolResult

                                {

                                    Success = true,

                                    Data = resultData

                                };

                            }

                            else

                            {

                                string errorMessage = responseObj["message"]?.Value<string>() ?? "获取display_content失败";

                                System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] ⚠️ API返回失败: {errorMessage}");

                                return new ToolResult { Success = false, Error = errorMessage };

                            }

                    }

                    else

                    {

                            string errorMessage = $"接口请求失败: {httpResult.StatusCode}";

                            try

                            {

                                var errorObj = JObject.Parse(responseContent);

                                errorMessage = errorObj["detail"]?.Value<string>() ?? errorMessage;

                            }

                            catch

                            {

                                // 如果解析失败，使用默认错误消息

                            }



                            System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] ⚠️ {errorMessage}");

                            return new ToolResult { Success = false, Error = errorMessage };

                    }

                }

                catch (Exception ex)

                {

                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] ⚠️ 异常: {ex.Message}");

                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_display_content] ⚠️ 堆栈跟踪: {ex.StackTrace}");

                    return new ToolResult { Success = false, Error = $"获取当前文档分块 display_content 失败: {ex.Message}" };

                }

            };
        }

    }

}


