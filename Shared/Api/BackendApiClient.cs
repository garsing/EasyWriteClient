using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// 业务后端 API（Api.BaseUrl）的统一 HTTP 封装：鉴权头、超时、401 刷新后重试。
    /// 不包含大模型流式请求（见 <see cref="LlmClient"/>）。
    /// </summary>
    public static class BackendApiClient
    {
        /// <summary>JSON 请求的通用结果（含失败时的响应体便于解析 detail）。</summary>
        public sealed class JsonResult
        {
            public bool Success { get; set; }
            public HttpStatusCode StatusCode { get; set; }
            public string Body { get; set; }
        }

        private static HttpClient CreateClient(TimeSpan? timeout = null)
        {
            var client = new HttpClient();
            client.Timeout = timeout ?? TimeSpan.FromSeconds(ConfigManager.Config.Api.TimeoutSeconds);
            return client;
        }

        private static void ApplyAuthHeaders(HttpClient client, UserService user)
        {
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", user.AccessToken);
            client.DefaultRequestHeaders.Remove("X-Username");
            client.DefaultRequestHeaders.Add("X-Username", user.UserName ?? "");
        }

        /// <summary>已登录用户：GET，401 时尝试刷新 token 后重试一次。</summary>
        /// <param name="authHandleAllErrors">为 true 时与历史侧栏一致：任意失败先走 <see cref="Auth.HandleApiResponse"/>（非 401 会抛错）。</param>
        public static async Task<JsonResult> GetAuthenticatedAsync(
            string url,
            bool retryOnUnauthorized = true,
            TimeSpan? timeout = null,
            bool authHandleAllErrors = false)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                using (var client = CreateClient(timeout))
                {
                    var user = UserService.Instance;
                    ApplyAuthHeaders(client, user);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var response = await client.GetAsync(url).ConfigureAwait(false);
                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (response.IsSuccessStatusCode)
                    {
                        return new JsonResult { Success = true, StatusCode = response.StatusCode, Body = body };
                    }

                    if (attempt == 0 && ShouldInvokeAuthOnFailure(retryOnUnauthorized, authHandleAllErrors, response.StatusCode))
                    {
                        var newToken = await Auth.Instance.HandleApiResponse(response, body).ConfigureAwait(false);
                        if (!string.IsNullOrEmpty(newToken))
                        {
                            continue;
                        }
                    }

                    return new JsonResult { Success = false, StatusCode = response.StatusCode, Body = body };
                }
            }

            return new JsonResult { Success = false, StatusCode = 0, Body = "" };
        }

        /// <summary>已登录用户：POST JSON，401 时尝试刷新 token 后重试一次。</summary>
        /// <param name="authHandleAllErrors">为 true 时与历史侧栏一致：任意失败先走 <see cref="Auth.HandleApiResponse"/>（非 401 会抛错）。</param>
        public static async Task<JsonResult> PostJsonAuthenticatedAsync(
            string url,
            string json,
            bool retryOnUnauthorized = true,
            TimeSpan? timeout = null,
            bool authHandleAllErrors = false)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                using (var client = CreateClient(timeout))
                {
                    var user = UserService.Instance;
                    ApplyAuthHeaders(client, user);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    using (var content = new StringContent(json ?? "{}", Encoding.UTF8, "application/json"))
                    {
                        var response = await client.PostAsync(url, content).ConfigureAwait(false);
                        var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (response.IsSuccessStatusCode)
                        {
                            return new JsonResult { Success = true, StatusCode = response.StatusCode, Body = body };
                        }

                        if (attempt == 0 && ShouldInvokeAuthOnFailure(retryOnUnauthorized, authHandleAllErrors, response.StatusCode))
                        {
                            var newToken = await Auth.Instance.HandleApiResponse(response, body).ConfigureAwait(false);
                            if (!string.IsNullOrEmpty(newToken))
                            {
                                continue;
                            }
                        }

                        return new JsonResult { Success = false, StatusCode = response.StatusCode, Body = body };
                    }
                }
            }

            return new JsonResult { Success = false, StatusCode = 0, Body = "" };
        }

        /// <summary>已登录用户：PUT JSON，401 时尝试刷新 token 后重试一次。</summary>
        public static async Task<JsonResult> PutJsonAuthenticatedAsync(
            string url,
            string json,
            bool retryOnUnauthorized = true,
            TimeSpan? timeout = null,
            bool authHandleAllErrors = false)
        {
            for (int attempt = 0; attempt < 2; attempt++)
            {
                using (var client = CreateClient(timeout))
                {
                    var user = UserService.Instance;
                    ApplyAuthHeaders(client, user);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    using (var content = new StringContent(json ?? "{}", Encoding.UTF8, "application/json"))
                    using (var request = new HttpRequestMessage(HttpMethod.Put, url) { Content = content })
                    {
                        var response = await client.SendAsync(request).ConfigureAwait(false);
                        var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (response.IsSuccessStatusCode)
                        {
                            return new JsonResult { Success = true, StatusCode = response.StatusCode, Body = body };
                        }

                        if (attempt == 0 && ShouldInvokeAuthOnFailure(retryOnUnauthorized, authHandleAllErrors, response.StatusCode))
                        {
                            var newToken = await Auth.Instance.HandleApiResponse(response, body).ConfigureAwait(false);
                            if (!string.IsNullOrEmpty(newToken))
                            {
                                continue;
                            }
                        }

                        return new JsonResult { Success = false, StatusCode = response.StatusCode, Body = body };
                    }
                }
            }

            return new JsonResult { Success = false, StatusCode = 0, Body = "" };
        }

        /// <summary>GET /users/me/settings</summary>
        public static Task<JsonResult> GetUserSettingsAsync()
        {
            string url = $"{ConfigManager.Config.Api.BaseUrl}/users/me/settings";
            return GetAuthenticatedAsync(url);
        }

        /// <summary>PUT /users/me/settings</summary>
        public static Task<JsonResult> PutUserSettingsAsync(string workspaceRoot)
        {
            string url = $"{ConfigManager.Config.Api.BaseUrl}/users/me/settings";
            string json = JsonConvert.SerializeObject(new { workspace_root = workspaceRoot });
            return PutJsonAuthenticatedAsync(url, json);
        }

        /// <summary>GET /users/me/llm-quota</summary>
        public static Task<JsonResult> GetLlmQuotaAsync()
        {
            string url = $"{ConfigManager.Config.Api.BaseUrl}/users/me/llm-quota";
            return GetAuthenticatedAsync(url);
        }

        private static bool ShouldInvokeAuthOnFailure(bool retryOnUnauthorized, bool authHandleAllErrors, HttpStatusCode statusCode)
        {
            if (authHandleAllErrors)
            {
                return true;
            }

            return retryOnUnauthorized && statusCode == HttpStatusCode.Unauthorized;
        }

        /// <summary>若已登录则带 Bearer，否则仅 JSON；不重试 401（与 MCP 工具列表同步行为一致）。</summary>
        public static async Task<JsonResult> PostJsonOptionalAuthAsync(string url, string json)
        {
            using (var client = CreateClient())
            {
                var user = UserService.Instance;
                if (user.IsLoggedIn && !string.IsNullOrEmpty(user.AccessToken))
                {
                    ApplyAuthHeaders(client, user);
                }

                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                using (var content = new StringContent(json ?? "{}", Encoding.UTF8, "application/json"))
                {
                    var response = await client.PostAsync(url, content).ConfigureAwait(false);
                    var body = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    return new JsonResult
                    {
                        Success = response.IsSuccessStatusCode,
                        StatusCode = response.StatusCode,
                        Body = body
                    };
                }
            }
        }

        /// <summary>GET <c>/files/download/{relativePath}</c> 并写入本地；404 返回 false。</summary>
        public static async Task<bool> DownloadFileFromUserDirectoryAsync(string relativePath, string localPath)
        {
            try
            {
                var userService = UserService.Instance;
                if (!userService.CheckLoginStatus())
                {
                    System.Diagnostics.Debug.WriteLine("[BackendApiClient] 用户未登录，无法从后端下载文件");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(relativePath))
                {
                    return false;
                }

                string baseUrl = ConfigManager.Config.Api.BaseUrl;
                string encodedPath = EncodeRelativePathForUrl(relativePath);
                string downloadUrl = $"{baseUrl}/files/download/{encodedPath}";

                for (int attempt = 0; attempt < 2; attempt++)
                {
                    using (var client = CreateClient())
                    {
                        ApplyAuthHeaders(client, userService);

                        System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 正在从后端下载文件: {downloadUrl}");

                        var response = await client.GetAsync(downloadUrl).ConfigureAwait(false);

                        if (response.IsSuccessStatusCode)
                        {
                            string directory = Path.GetDirectoryName(localPath);
                            if (!string.IsNullOrEmpty(directory))
                            {
                                Directory.CreateDirectory(directory);
                            }

                            using (var fileStream = File.Create(localPath))
                            {
                                await response.Content.CopyToAsync(fileStream).ConfigureAwait(false);
                            }

                            System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 成功下载文件到本地: {localPath}");
                            return true;
                        }

                        if (response.StatusCode == HttpStatusCode.NotFound)
                        {
                            System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 后端文件不存在: {relativePath}");
                            return false;
                        }

                        var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (attempt == 0 && response.StatusCode == HttpStatusCode.Unauthorized)
                        {
                            var newToken = await Auth.Instance.HandleApiResponse(response, responseContent).ConfigureAwait(false);
                            if (!string.IsNullOrEmpty(newToken))
                            {
                                continue;
                            }
                        }

                        System.Diagnostics.Debug.WriteLine(
                            $"[BackendApiClient] 下载文件HTTP错误: {response.StatusCode} - {responseContent}");
                        return false;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 下载文件异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>POST <c>/files/upload</c>，multipart：<c>file</c> + <c>relative_path</c>。</summary>
        public static async Task<bool> UploadUserFileAsync(string filePath, string relativePath)
        {
            try
            {
                var userService = UserService.Instance;
                if (!userService.CheckLoginStatus())
                {
                    System.Diagnostics.Debug.WriteLine("[BackendApiClient] 用户未登录，跳过文件上传");
                    return false;
                }

                if (!File.Exists(filePath))
                {
                    System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 文件不存在: {filePath}");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(relativePath))
                {
                    System.Diagnostics.Debug.WriteLine("[BackendApiClient] relative_path 不能为空");
                    return false;
                }

                string baseUrl = ConfigManager.Config.Api.BaseUrl;
                string uploadUrl = $"{baseUrl}/files/upload";
                string uploadFilename = Path.GetFileName(filePath);

                for (int attempt = 0; attempt < 2; attempt++)
                {
                    using (var client = CreateClient())
                    {
                        ApplyAuthHeaders(client, userService);

                        using (var formData = new MultipartFormDataContent())
                        using (var fileStream = File.OpenRead(filePath))
                        using (var streamContent = new StreamContent(fileStream))
                        {
                            streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                            formData.Add(streamContent, "file", uploadFilename);
                            formData.Add(new StringContent(relativePath), "relative_path");

                            System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 正在上传文件到: {uploadUrl} ({relativePath})");

                            var response = await client.PostAsync(uploadUrl, formData).ConfigureAwait(false);
                            var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                            if (response.IsSuccessStatusCode)
                            {
                                var uploadResponse = JsonConvert.DeserializeObject<UploadResponseDto>(responseContent);
                                if (uploadResponse != null && uploadResponse.success)
                                {
                                    System.Diagnostics.Debug.WriteLine(
                                        $"[BackendApiClient] 文件上传成功: {uploadResponse.data?.filename ?? uploadFilename}");
                                    return true;
                                }

                                System.Diagnostics.Debug.WriteLine(
                                    $"[BackendApiClient] 文件上传失败: {uploadResponse?.message ?? "未知错误"}");
                                return false;
                            }

                            if (attempt == 0 && response.StatusCode == HttpStatusCode.Unauthorized)
                            {
                                var newToken = await Auth.Instance.HandleApiResponse(response, responseContent).ConfigureAwait(false);
                                if (!string.IsNullOrEmpty(newToken))
                                {
                                    continue;
                                }
                            }

                            System.Diagnostics.Debug.WriteLine(
                                $"[BackendApiClient] 文件上传HTTP错误: {response.StatusCode} - {responseContent}");
                            return false;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 文件上传异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>POST <c>/curr_doc/process</c>，multipart：文件 + <c>doc_uuid</c>；超时 300 秒。</summary>
        public static async Task<JsonResult> PostProcessDocumentMultipartAsync(string filePath, string docUuid)
        {
            var userService = UserService.Instance;
            if (!userService.CheckLoginStatus())
            {
                return new JsonResult { Success = false, StatusCode = 0, Body = "" };
            }

            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                return new JsonResult { Success = false, StatusCode = 0, Body = "" };
            }

            string baseUrl = ConfigManager.Config.Api.BaseUrl;
            string apiUrl = $"{baseUrl}/curr_doc/process";

            for (int attempt = 0; attempt < 2; attempt++)
            {
                using (var client = CreateClient(TimeSpan.FromSeconds(300)))
                {
                    ApplyAuthHeaders(client, userService);

                    using (var formData = new MultipartFormDataContent())
                    using (var fileStream = File.OpenRead(filePath))
                    using (var streamContent = new StreamContent(fileStream))
                    {
                        streamContent.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
                        string fileName = Path.GetFileName(filePath);
                        formData.Add(streamContent, "file", fileName);
                        formData.Add(new StringContent(docUuid), "doc_uuid");

                        System.Diagnostics.Debug.WriteLine($"[BackendApiClient] 正在上传文档到: {apiUrl}");

                        var response = await client.PostAsync(apiUrl, formData).ConfigureAwait(false);
                        var responseContent = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (response.IsSuccessStatusCode)
                        {
                            return new JsonResult { Success = true, StatusCode = response.StatusCode, Body = responseContent };
                        }

                        if (attempt == 0 && response.StatusCode == HttpStatusCode.Unauthorized)
                        {
                            var newToken = await Auth.Instance.HandleApiResponse(response, responseContent).ConfigureAwait(false);
                            if (!string.IsNullOrEmpty(newToken))
                            {
                                continue;
                            }
                        }

                        return new JsonResult { Success = false, StatusCode = response.StatusCode, Body = responseContent };
                    }
                }
            }

            return new JsonResult { Success = false, StatusCode = 0, Body = "" };
        }

        private static string EncodeRelativePathForUrl(string relativePath)
        {
            var segments = relativePath.Replace('\\', '/').Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join("/", segments.Select(Uri.EscapeDataString));
        }

        private sealed class UploadResponseDto
        {
            public bool success { get; set; }
            public string message { get; set; }
            public UploadDataDto data { get; set; }
        }

        private sealed class UploadDataDto
        {
            public string filename { get; set; }
        }
    }
}
