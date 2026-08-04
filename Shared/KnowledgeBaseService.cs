using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    /// <summary>
    /// 知识库 API 服务类
    /// 仅保留 GetApiConfig 方法，用于前端获取 API 配置
    /// 其他 API 调用已改为前端直接调用
    /// </summary>
    public class KnowledgeBaseService
    {
        private static readonly string BaseUrl = ConfigManager.Config.Api.BaseUrl;
        private static readonly int TimeoutSeconds = ConfigManager.Config.Api.TimeoutSeconds;

        /// <summary>
        /// 创建知识库 API 请求头
        /// </summary>
        private static Dictionary<string, string> CreateApiHeaders()
        {
            var headers = new Dictionary<string, string>();

            // 检查用户是否已登录
            if (UserService.Instance.CheckLoginStatus())
            {
                var userService = UserService.Instance;

                // 添加认证令牌
                if (!string.IsNullOrEmpty(userService.AccessToken))
                {
                    headers["Authorization"] = $"Bearer {userService.AccessToken}";
                }

                // 添加用户名
                if (!string.IsNullOrEmpty(userService.UserName))
                {
                    headers["X-Username"] = userService.UserName;
                }
            }

            return headers;
        }

        /// <summary>
        /// 获取 API 配置信息（用于前端直接调用 API）。
        /// 未登录也返回 baseUrl；headers 仅在已登录时含 Authorization / X-Username。
        /// </summary>
        public static object GetApiConfig(object data)
        {
            try
            {
                var headers = CreateApiHeaders();
                var config = new
                {
                    success = true,
                    data = new
                    {
                        baseUrl = BaseUrl,
                        headers = headers
                    }
                };

                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseService] 返回 API 配置: BaseUrl={BaseUrl}, loggedIn={UserService.Instance.CheckLoginStatus()}");
                return config;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseService] 获取 API 配置失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        /// <summary>
        /// 对话附件：由宿主拦截拖放后的本地路径，直接 multipart 上传（与前端 <c>uploadDocumentForChat</c> 一致）。
        /// </summary>
        public static async Task<HostChatUploadResult> UploadDocumentForChatFromPathAsync(string localPath)
        {
            var result = new HostChatUploadResult();
            try
            {
                if (string.IsNullOrWhiteSpace(localPath) || !File.Exists(localPath))
                {
                    result.Message = "文件不存在或路径无效";
                    return result;
                }

                if (!UserService.Instance.CheckLoginStatus())
                {
                    result.Message = "请先登录";
                    return result;
                }

                string baseUrl = ConfigManager.Config.Api.BaseUrl?.TrimEnd('/') ?? "";
                if (string.IsNullOrEmpty(baseUrl))
                {
                    result.Message = "未配置 API 地址";
                    return result;
                }

                string uploadUrl = $"{baseUrl}/knowledge/document/upload";
                var headers = CreateApiHeaders();

                using (var client = new HttpClient())
                {
                    client.Timeout = TimeSpan.FromSeconds(Math.Max(TimeoutSeconds, 300));
                    foreach (var kv in headers)
                        client.DefaultRequestHeaders.TryAddWithoutValidation(kv.Key, kv.Value);

                    using (var form = new MultipartFormDataContent())
                    {
                        var fileStream = File.OpenRead(localPath);
                        var streamContent = new StreamContent(fileStream);
                        streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        form.Add(streamContent, "file", Path.GetFileName(localPath));
                        form.Add(new StringContent("default_upload"), "knowledge_base_uuid");

                        var response = await client.PostAsync(uploadUrl, form).ConfigureAwait(false);
                        string body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        JObject jo;
                        try
                        {
                            jo = JObject.Parse(body);
                        }
                        catch (JsonReaderException)
                        {
                            result.Message = "上传响应不是有效 JSON";
                            return result;
                        }

                        bool ok = jo["success"]?.Value<bool>() ?? false;
                        if (!response.IsSuccessStatusCode || !ok)
                        {
                            result.Message = jo["message"]?.ToString() ?? jo["detail"]?.ToString() ?? "上传失败";
                            return result;
                        }

                        result.StorageDocUuid = jo["data"]?["storage_doc_uuid"]?.ToString();
                        if (string.IsNullOrEmpty(result.StorageDocUuid))
                        {
                            result.Message = "上传响应缺少 storage_doc_uuid";
                            return result;
                        }

                        var processingStatus = jo["data"]?["processing_status"]?.Value<int>() ?? 0;
                        result.SkipProcess = processingStatus == 2;
                        result.Success = true;
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseService] UploadDocumentForChatFromPathAsync: {ex.Message}");
                result.Message = ex.Message;
                return result;
            }
        }
    }

    /// <summary>宿主侧对话文档上传结果（供 WebView2 拖放兜底）</summary>
    public sealed class HostChatUploadResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public string StorageDocUuid { get; set; }
        /// <summary>index 已处理完成（processing_status=2）时跳过 process，直接 by-document。</summary>
        public bool SkipProcess { get; set; }
    }
}
