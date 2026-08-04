using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;

namespace WordAddIn1
{
    /// <summary>
    /// 认证管理器，负责令牌刷新、错误处理和自动重新登录
    /// </summary>
    public class Auth
    {
        private static Auth _instance;
        private static readonly object _lock = new object();
        private bool _isRefreshing = false;
        private readonly Queue<TaskCompletionSource<string>> _refreshQueue = new Queue<TaskCompletionSource<string>>();

        /// <summary>
        /// 获取Auth单例实例
        /// </summary>
        public static Auth Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new Auth();
                        }
                    }
                }
                return _instance;
            }
        }

        private Auth()
        {
            // 私有构造函数，确保单例模式
        }

        /// <summary>
        /// 处理HTTP响应，根据状态码和错误信息进行相应处理
        /// </summary>
        /// <param name="response">HTTP响应</param>
        /// <param name="responseContent">响应内容</param>
        /// <returns>处理后的新Access Token，如果不需要刷新则返回null</returns>
        public async Task<string> HandleApiResponse(HttpResponseMessage response, string responseContent)
        {
            if (response.IsSuccessStatusCode)
            {
                return null; // 成功响应，不需要处理
            }

            if (response.StatusCode == System.Net.HttpStatusCode.Unauthorized)
            {
                return await HandleUnauthorizedResponse(response, responseContent);
            }

            // 其他错误直接抛出异常
            throw new LlmException($"API请求失败: HTTP {response.StatusCode} - {responseContent}");
        }

        /// <summary>
        /// 处理401未授权响应
        /// </summary>
        private async Task<string> HandleUnauthorizedResponse(HttpResponseMessage response, string responseContent)
        {
            try
            {
                // 尝试解析错误详情
                var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseContent);
                string errorDetail = errorResponse?.detail ?? "未知认证错误";

                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 收到401错误: {errorDetail}");

                // 根据错误详情判断处理方式
                if (IsTokenExpiredError(errorDetail))
                {
                    // Access Token过期，尝试刷新
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] Access Token过期，尝试刷新");
                    return await RefreshAccessToken();
                }
                else if (IsInvalidTokenError(errorDetail))
                {
                    // Refresh Token无效或过期，需要重新登录
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] Refresh Token无效，需要重新登录");
                    await TriggerReLogin("认证信息已失效，请重新登录");
                    throw new LlmException("认证失效，请重新登录");
                }
                else
                {
                    // 其他认证错误
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 其他认证错误: {errorDetail}");
                    await TriggerReLogin($"认证失败: {errorDetail}");
                    throw new LlmException($"认证失败: {errorDetail}");
                }
            }
            catch (JsonException)
            {
                // JSON解析失败，当作一般错误处理
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 无法解析错误响应: {responseContent}");
                await TriggerReLogin("认证服务异常，请重新登录");
                throw new LlmException("认证服务异常，请重新登录");
            }
        }

        /// <summary>
        /// 判断是否为Access Token过期错误
        /// </summary>
        private bool IsTokenExpiredError(string errorDetail)
        {
            return errorDetail.Contains("访问令牌已过期") ||
                   errorDetail.Contains("令牌已过期") ||
                   errorDetail.Contains("token expired");
        }

        /// <summary>
        /// 判断是否为Refresh Token无效错误
        /// </summary>
        private bool IsInvalidTokenError(string errorDetail)
        {
            return errorDetail.Contains("无效的访问令牌") ||
                   errorDetail.Contains("无效或过期的刷新令牌") ||
                   errorDetail.Contains("刷新令牌已过期") ||
                   errorDetail.Contains("刷新令牌不匹配") ||
                   errorDetail.Contains("用户名验证失败") ||
                   errorDetail.Contains("令牌格式错误") ||
                   errorDetail.Contains("令牌类型错误") ||
                   errorDetail.Contains("缺少") && errorDetail.Contains("请求头");
        }

        /// <summary>
        /// 刷新Access Token
        /// </summary>
        /// <returns>新的Access Token</returns>
        public async Task<string> RefreshAccessToken()
        {
            EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] === 开始RefreshAccessToken方法 ===");

            if (_isRefreshing)
            {
                // 如果正在刷新，加入队列等待
                EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] 正在刷新中，加入等待队列");
                var tcs = new TaskCompletionSource<string>();
                _refreshQueue.Enqueue(tcs);
                return await tcs.Task;
            }

            _isRefreshing = true;
            EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] 开始刷新Access Token");

            try
            {
                var userService = UserService.Instance;
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] UserService状态 - RefreshToken: {(string.IsNullOrEmpty(userService.RefreshToken) ? "null/empty" : "存在")}, UserName: {userService.UserName ?? "null"}");

                if (string.IsNullOrEmpty(userService.RefreshToken) ||
                    string.IsNullOrEmpty(userService.UserName))
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] 缺少Refresh Token或用户名，触发重新登录");
                    await TriggerReLogin("缺少认证信息，请重新登录");
                    throw new LlmException("缺少认证信息，请重新登录");
                }

                using (var httpClient = new HttpClient())
                {
                    // 设置超时时间
                    httpClient.Timeout = TimeSpan.FromSeconds(ConfigManager.Config.Api.TimeoutSeconds);
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 配置信息 - BaseUrl: {ConfigManager.Config.Api.BaseUrl}, Timeout: {ConfigManager.Config.Api.TimeoutSeconds}");

                    // 构建请求头
                    httpClient.DefaultRequestHeaders.Add("X-Refresh-Token", userService.RefreshToken);
                    httpClient.DefaultRequestHeaders.Add("X-Username", userService.UserName);
                    httpClient.DefaultRequestHeaders.Add("Accept", "application/json");
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 请求头 - X-Refresh-Token: {(string.IsNullOrEmpty(userService.RefreshToken) ? "null" : "存在")}, X-Username: {userService.UserName}");

                    // 检查Endpoints字典
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] Endpoints字典包含的键: {string.Join(", ", ConfigManager.Config.Api.Endpoints.Keys)}");

                    // 检查Refresh端点是否存在，如果不存在则重新加载配置
                    if (!ConfigManager.Config.Api.Endpoints.ContainsKey("Refresh"))
                    {
                        EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] Refresh端点不存在，尝试重新加载配置");
                        string configPath = ConfigManager.GetConfigFilePath();
                        EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 配置文件路径: {configPath}");
                        if (File.Exists(configPath))
                        {
                            string configContent = File.ReadAllText(configPath);
                            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 配置文件内容预览: {configContent.Substring(0, Math.Min(200, configContent.Length))}");
                        }
                        else
                        {
                            EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] 配置文件不存在！");
                        }
                        ConfigManager.ReloadConfig();
                        EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 重新加载后Endpoints字典包含的键: {string.Join(", ", ConfigManager.Config.Api.Endpoints.Keys)}");
                    }

                    // 发送刷新请求
                    string refreshUrl = $"{ConfigManager.Config.Api.BaseUrl}{ConfigManager.Config.Api.Endpoints["Refresh"]}";
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 发送刷新请求到: {refreshUrl}");

                    using (var response = await httpClient.PostAsync(refreshUrl, null))
                    {
                        string responseContent = await response.Content.ReadAsStringAsync();
                        EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 刷新响应状态码: {(int)response.StatusCode}");

                        if (response.IsSuccessStatusCode)
                        {
                            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 刷新响应内容: {responseContent}");

                            // 解析响应
                            var refreshResponse = JsonConvert.DeserializeObject<RefreshResponse>(responseContent);
                            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 解析结果 - access_token: {(refreshResponse?.access_token != null ? "存在" : "null")}, refresh_token: {(refreshResponse?.refresh_token != null ? "存在" : "null")}, expires_in: {refreshResponse?.expires_in}");
                            if (refreshResponse != null &&
                                !string.IsNullOrEmpty(refreshResponse.access_token) &&
                                !string.IsNullOrEmpty(refreshResponse.refresh_token))
                            {
                                // 计算过期时间：优先使用expires_at字段，否则使用expires_in
                                DateTime? tokenExpiry = null;
                                if (!string.IsNullOrEmpty(refreshResponse.expires_at))
                                {
                                    // 解析CST格式的时间戳 (如: "2025-12-15T20:10:35+08:00")
                                    if (DateTime.TryParse(refreshResponse.expires_at, out DateTime parsedExpiry))
                                    {
                                        tokenExpiry = parsedExpiry;
                                        EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 使用服务器提供的过期时间: {tokenExpiry}");
                                    }
                                }

                                if (tokenExpiry == null && refreshResponse.expires_in > 0)
                                {
                                    // 备用方案：使用expires_in计算
                                    tokenExpiry = DateTime.Now.AddSeconds(refreshResponse.expires_in);
                                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 使用expires_in计算过期时间: {tokenExpiry}");
                                }

                                // 更新UserService中的令牌
                                userService.UpdateTokens(
                                    refreshResponse.access_token,
                                    refreshResponse.refresh_token,
                                    tokenExpiry
                                );

                                EasyWriteDiagnostics.Log(DebugCategory.Auth, "[AUTH] Access Token刷新成功");

                                // 处理等待队列
                                while (_refreshQueue.Count > 0)
                                {
                                    var waitingTask = _refreshQueue.Dequeue();
                                    waitingTask.SetResult(refreshResponse.access_token);
                                }

                                return refreshResponse.access_token;
                            }
                            else
                            {
                                throw new LlmException("刷新响应格式错误");
                            }
                        }
                        else
                        {
                            // 刷新失败
                            string errorMsg = $"刷新失败: HTTP {(int)response.StatusCode}";
                            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 刷新失败 - 状态码: {(int)response.StatusCode}, 响应内容: {responseContent}");

                            if (!string.IsNullOrEmpty(responseContent))
                            {
                                try
                                {
                                    var errorResponse = JsonConvert.DeserializeObject<ErrorResponse>(responseContent);
                                    errorMsg += $" - {errorResponse?.detail ?? responseContent}";
                                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 解析错误响应成功 - detail: {errorResponse?.detail}");
                                }
                                catch (Exception ex)
                                {
                                    errorMsg += $" - {responseContent}";
                                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 解析错误响应失败: {ex.Message}");
                                }
                            }

                            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] {errorMsg}");

                            // 刷新失败，触发重新登录
                            await TriggerReLogin("令牌刷新失败，请重新登录");
                            throw new LlmException($"认证失败: {errorMsg}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 刷新过程中发生异常: {ex.Message}");
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 异常类型: {ex.GetType().Name}");
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 异常堆栈: {ex.StackTrace}");

                // 处理等待队列的错误
                while (_refreshQueue.Count > 0)
                {
                    var waitingTask = _refreshQueue.Dequeue();
                    waitingTask.SetException(ex);
                }

                // 触发重新登录
                await TriggerReLogin($"令牌刷新异常: {ex.Message}");
                throw new LlmException($"认证失败: 令牌刷新异常: {ex.Message}", ex);
            }
            finally
            {
                _isRefreshing = false;
            }
        }

        /// <summary>
        /// 触发重新登录，弹出登录面板
        /// </summary>
        /// <param name="message">显示给用户的消息</param>
        private async Task TriggerReLogin(string message)
        {
            EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 触发重新登录: {message}");

            // 在UI线程上显示登录面板
            if (Application.OpenForms.Count > 0)
            {
                var mainForm = Application.OpenForms[0];
                if (mainForm.InvokeRequired)
                {
                    await (Task)mainForm.Invoke(new Action(() => ShowLoginDialog(message)));
                }
                else
                {
                    ShowLoginDialog(message);
                }
            }
        }

        /// <summary>
        /// 显示登录对话框
        /// </summary>
        private void ShowLoginDialog(string message)
        {
            try
            {
                UserService.Instance.Logout();

                if (!LoginForm.IsOpen && !UserSettingsForm.IsOpen)
                {
                    MessageBox.Show(message, "需要重新登录", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 登录/设置窗口已打开，跳过 MessageBox: {message}");
                }

                var owner = Application.OpenForms.Count > 0 ? Application.OpenForms[0] : null;
                _ = LoginForm.ShowDialogAsync(owner, logoutFirst: false);

                OnReLoginRequired?.Invoke(this, new ReLoginEventArgs(message));
            }
            catch (Exception ex)
            {
                EasyWriteDiagnostics.Log(DebugCategory.Auth, $"[AUTH] 显示登录对话框时发生错误: {ex.Message}");
            }
        }

        /// <summary>
        /// 重新登录需要事件
        /// </summary>
        public event EventHandler<ReLoginEventArgs> OnReLoginRequired;

        /// <summary>
        /// 检查是否需要重新登录
        /// </summary>
        /// <returns>true如果需要重新登录</returns>
        public bool ShouldReLogin()
        {
            var userService = UserService.Instance;
            return !userService.CheckLoginStatus();
        }
    }

    /// <summary>
    /// 错误响应模型
    /// </summary>
    public class ErrorResponse
    {
        public string detail { get; set; }
    }

    /// <summary>
    /// 刷新令牌响应模型
    /// </summary>
    public class RefreshResponse
    {
        public string access_token { get; set; }
        public string refresh_token { get; set; }
        public string token_type { get; set; }
        public int expires_in { get; set; }
        public string expires_at { get; set; }
    }

    /// <summary>
    /// 重新登录事件参数
    /// </summary>
    public class ReLoginEventArgs : EventArgs
    {
        public string Message { get; }

        public ReLoginEventArgs(string message)
        {
            Message = message;
        }
    }
}
