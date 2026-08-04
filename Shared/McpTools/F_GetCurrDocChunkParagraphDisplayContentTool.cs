using System;

using System.Collections.Generic;

using System.Threading.Tasks;

using Newtonsoft.Json;

using Newtonsoft.Json.Linq;

using Word = Microsoft.Office.Interop.Word;



namespace WordAddIn1

{

    /// <summary>

    /// 获取当前文档指定分块的 paragraph_display_content 工具

    /// 根据 chunk_idx 从 API 获取当前文档（DocumentState）对应分块的 paragraph_display_content

    /// </summary>

    public static class F_GetCurrDocChunkParagraphDisplayContentTool

    {

        /// <summary>

        /// 注册获取当前文档分块 paragraph_display_content 工具

        /// </summary>

        public static void Register(

            Dictionary<string, Func<Dictionary<string, object>, Task<ToolResult>>> toolRegistry,

            object wordApplication)

        {
            toolRegistry["F_get_curr_doc_chunk_paragraph_display_content"] = async (args) =>

            {

                try

                {

                    if (!ChannelDocument.TryResolve(args, wordApplication, out Word.Document document, out ToolResult resolveError))

                    {

                        return resolveError;

                    }



                    if (!args.ContainsKey("chunk_idx"))

                    {

                        return new ToolResult { Success = false, Error = "缺少必需参数: chunk_idx" };

                    }



                    object chunkIdxObj = args["chunk_idx"];

                    int chunkIdx;



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



                    if (chunkIdx < 1)

                    {

                        return new ToolResult { Success = false, Error = "chunk_idx必须大于等于1" };

                    }



                    string docUuid = DocumentState.CurrDocUuid;

                    int snapshotIdx = DocumentState.SnapshotIdx;



                    if (string.IsNullOrEmpty(docUuid))

                    {

                        return new ToolResult { Success = false, Error = "未找到文档UUID，请先处理文档" };

                    }



                    if (snapshotIdx < 1)

                    {

                        var paraDisplays = DocumentState.ChunkParagraphDisplayContents;

                        if (paraDisplays != null && chunkIdx >= 1 && chunkIdx <= paraDisplays.Count)

                        {

                            string localParagraphDisplayContent = paraDisplays[chunkIdx - 1] ?? "";

                            var localResultData = new Dictionary<string, object>

                            {

                                { "paragraph_display_content", localParagraphDisplayContent },

                                { "chunk_idx", chunkIdx },

                                { "doc_uuid", docUuid },

                                { "snapshot_idx", snapshotIdx },

                                { "source", "local" }

                            };



                            return new ToolResult

                            {

                                Success = true,

                                Data = localResultData

                            };

                        }



                        return new ToolResult { Success = false, Error = "快照索引无效，请先处理文档" };

                    }



                    var userService = UserService.Instance;

                    if (!userService.CheckLoginStatus())

                    {

                        return new ToolResult { Success = false, Error = "用户未登录，无法调用接口" };

                    }



                    string baseUrl = ConfigManager.Config.Api.BaseUrl;

                    string apiUrl = $"{baseUrl}/curr_doc/paragraph_display_content?doc_uuid={Uri.EscapeDataString(docUuid)}&snapshot_idx={snapshotIdx}&chunk_idx={chunkIdx}";



                    System.Diagnostics.Debug.WriteLine($"[F_get_curr_doc_chunk_paragraph_display_content] 调用API: {apiUrl}");



                    var httpResult = await BackendApiClient.GetAuthenticatedAsync(

                        apiUrl,

                        retryOnUnauthorized: true,

                        TimeSpan.FromSeconds(30));

                    var responseContent = httpResult.Body;



                    if (httpResult.Success)

                    {

                            var responseObj = JObject.Parse(responseContent);

                            if (responseObj["success"]?.Value<bool>() == true && responseObj["data"] != null)

                            {

                                var data = responseObj["data"];

                                string paragraphDisplayContent = data["paragraph_display_content"]?.Value<string>() ?? "";



                                var resultData = new Dictionary<string, object>

                                {

                                    { "paragraph_display_content", paragraphDisplayContent },

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

                                string errorMessage = responseObj["message"]?.Value<string>() ?? "获取paragraph_display_content失败";

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

                            }



                            return new ToolResult { Success = false, Error = errorMessage };

                    }

                }

                catch (Exception ex)

                {

                    return new ToolResult { Success = false, Error = $"获取当前文档分块 paragraph_display_content 失败: {ex.Message}" };

                }

            };
        }

    }

}
