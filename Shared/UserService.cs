using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace WordAddIn1
{
    /// <summary>
    /// 用户服务类 - 全局用户状态管理
    /// </summary>
    public class UserService
    {
        private static UserService _instance;
        private static readonly object _lock = new object();
        private readonly object _stateLock = new object();
        private bool _isLoggingOut;

        // API配置 - 从统一配置文件读取
        private static readonly HttpClient _httpClient;

        // 配置属性
        private static string BaseUrl => ConfigManager.Config.Api.BaseUrl;
        private static int ApiTimeoutSeconds => ConfigManager.Config.Api.TimeoutSeconds;
        private static int ApiMaxRetries => ConfigManager.Config.Api.MaxRetries;
        private static string LoginEndpoint => ConfigManager.Config.Api.Endpoints["Login"];
        private static string RegisterEndpoint => ConfigManager.Config.Api.Endpoints["Register"];
        private static string LoginSmsEndpoint =>
            ConfigManager.Config.Api.Endpoints.TryGetValue("LoginSms", out var loginSms)
                ? loginSms
                : "/auth/login/sms";
        private static string SmsSendEndpoint =>
            ConfigManager.Config.Api.Endpoints.TryGetValue("SmsSend", out var smsSend)
                ? smsSend
                : "/auth/sms/send";

        // 静态构造函数 - 初始化HttpClient
        static UserService()
        {
            _httpClient = new HttpClient();
            _httpClient.Timeout = TimeSpan.FromSeconds(ApiTimeoutSeconds);
            _httpClient.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
        }

        // 用户状态属性
        public bool IsLoggedIn { get; private set; }
        public string UserName { get; private set; }
        public string UserId { get; private set; }
        public string AccessToken { get; private set; }
        public string RefreshToken { get; private set; }
        public DateTime? LoginTime { get; private set; }
        public DateTime? TokenExpiry { get; private set; }

        /// <summary>MySQL 原始配置值（设置页显示用）。</summary>
        public string WorkspaceRootConfigured { get; private set; }

        /// <summary>D33 remap 后的本机有效路径（F_* 读写用）。</summary>
        public string WorkspaceRootEffective { get; private set; }

        /// <summary>本机无可用盘符且用户尚未选择目录时为 true。</summary>
        public bool WorkspaceDirectorySelectionRequired { get; private set; }

        // API响应模型
        private class ApiResponse<T>
        {
            public bool success { get; set; }
            public string message { get; set; }
            public T data { get; set; }
        }

        private class LoginResponse
        {
            public string access_token { get; set; }
            public string refresh_token { get; set; }
            public string token_type { get; set; }
            public int expires_in { get; set; }
            public UserData user { get; set; }
        }

        private class UserData
        {
            public int id { get; set; }
            public string username { get; set; }
            public string email { get; set; }
            public string phone { get; set; }
            public string avatar { get; set; }
            public string status { get; set; }
            public bool email_verified { get; set; }
            public bool phone_verified { get; set; }
            public DateTime created_at { get; set; }
            public DateTime? updated_at { get; set; }
            public DateTime? last_login_at { get; set; }
        }

        private class RegisterRequest
        {
            public string phone { get; set; }
            public string sms_code { get; set; }
            public string password { get; set; }
        }

        private class LoginRequest
        {
            public string credential { get; set; }
            public string password { get; set; }
        }

        private class SmsLoginRequest
        {
            public string phone { get; set; }
            public string sms_code { get; set; }
        }

        private class SmsSendRequest
        {
            public string phone { get; set; }
            public string scene { get; set; }
        }

        private class UserSettingsData
        {
            public string workspace_root { get; set; }
        }

        private class LlmQuotaData
        {
            public string plan_code { get; set; }
            public string plan_name { get; set; }
            public string provider { get; set; }
            public string base_model { get; set; }
            public string usage_date { get; set; }
            public double daily_quota_inspiration { get; set; }
            public double daily_used_inspiration { get; set; }
            public double daily_used_cny { get; set; }
            public double inspiration_per_cny { get; set; }
            public double loose_inspiration { get; set; }
            public string quota_period { get; set; }
            public string plan_expires_at { get; set; }
            public double? period_quota_inspiration { get; set; }
            public double? period_used_inspiration { get; set; }
        }

        // 私有构造函数
        private UserService()
        {
            // 初始化时从设置加载状态
            LoadFromSettings();
        }

        /// <summary>
        /// 获取单例实例
        /// </summary>
        public static UserService Instance
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        if (_instance == null)
                        {
                            _instance = new UserService();
                        }
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// 用户登录
        /// </summary>
        public async Task<bool> LoginAsync(string credential, string password)
        {
            var result = await LoginWithMessageAsync(credential, password).ConfigureAwait(false);
            return result.Success;
        }

        /// <summary>
        /// 用户登录（含错误文案）
        /// </summary>
        public async Task<RegisterResult> LoginWithMessageAsync(string credential, string password)
        {
            try
            {
                return await ExecuteWithRetry(async () =>
                {
                    var loginRequest = new LoginRequest
                    {
                        credential = credential,
                        password = password
                    };

                    var json = JsonConvert.SerializeObject(loginRequest);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await _httpClient.PostAsync($"{BaseUrl}{LoginEndpoint}", content);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    System.Diagnostics.Debug.WriteLine($"登录API响应 - 状态码: {response.StatusCode}");

                    if (!response.IsSuccessStatusCode)
                    {
                        var detail = TryExtractErrorDetail(responseContent);
                        return RegisterResult.Fail(detail ?? "手机号或密码错误");
                    }

                    if (!TryApplyTokenResponse(responseContent))
                    {
                        return RegisterResult.Fail("手机号或密码错误");
                    }

                    return RegisterResult.Ok();
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] LoginWithMessageAsync 异常: {ex.Message}");
                return RegisterResult.Fail("登录失败，请稍后重试");
            }
        }

        /// <summary>
        /// 验证码登录
        /// </summary>
        public async Task<RegisterResult> LoginBySmsWithMessageAsync(string phone, string smsCode)
        {
            try
            {
                return await ExecuteWithRetry(async () =>
                {
                    var request = new SmsLoginRequest
                    {
                        phone = phone,
                        sms_code = smsCode
                    };

                    var json = JsonConvert.SerializeObject(request);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await _httpClient.PostAsync($"{BaseUrl}{LoginSmsEndpoint}", content);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        var detail = TryExtractErrorDetail(responseContent);
                        return RegisterResult.Fail(detail ?? "验证码错误或已过期");
                    }

                    if (!TryApplyTokenResponse(responseContent))
                    {
                        return RegisterResult.Fail("验证码错误或已过期");
                    }

                    return RegisterResult.Ok();
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] LoginBySmsWithMessageAsync 异常: {ex.Message}");
                return RegisterResult.Fail("登录失败，请稍后重试");
            }
        }

        /// <summary>
        /// 发送短信验证码
        /// </summary>
        public async Task<RegisterResult> SendSmsCodeAsync(string phone, string scene)
        {
            try
            {
                return await ExecuteWithRetry(async () =>
                {
                    var request = new SmsSendRequest
                    {
                        phone = phone,
                        scene = scene
                    };

                    var json = JsonConvert.SerializeObject(request);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await _httpClient.PostAsync($"{BaseUrl}{SmsSendEndpoint}", content);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    if (!response.IsSuccessStatusCode)
                    {
                        var detail = TryExtractErrorDetail(responseContent);
                        return RegisterResult.Fail(detail ?? "验证码发送失败");
                    }

                    var apiResponse = JsonConvert.DeserializeObject<ApiResponse<object>>(responseContent);
                    if (apiResponse != null && apiResponse.success)
                    {
                        return RegisterResult.Ok(apiResponse.message ?? "验证码已发送");
                    }

                    return RegisterResult.Fail(apiResponse?.message ?? "验证码发送失败");
                }).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] SendSmsCodeAsync 异常: {ex.Message}");
                return RegisterResult.Fail("验证码发送失败，请稍后重试");
            }
        }

        private bool TryApplyTokenResponse(string responseContent)
        {
            var tokenResponse = JsonConvert.DeserializeObject<LoginResponse>(responseContent);
            if (tokenResponse == null || tokenResponse.user == null || string.IsNullOrEmpty(tokenResponse.access_token))
            {
                System.Diagnostics.Debug.WriteLine($"登录失败，无法获取有效的登录信息: {responseContent}");
                return false;
            }

            IsLoggedIn = true;
            UserName = tokenResponse.user.username;
            UserId = tokenResponse.user.id.ToString();
            AccessToken = tokenResponse.access_token;
            RefreshToken = tokenResponse.refresh_token;
            LoginTime = DateTime.Now;

            if (tokenResponse.expires_in > 0)
            {
                TokenExpiry = DateTime.Now.AddSeconds(tokenResponse.expires_in);
            }

            SaveToSettings();
            OnUserLoggedIn?.Invoke(this, new UserEventArgs(UserName, UserId));
            System.Diagnostics.Debug.WriteLine($"登录成功: {UserName}, Token过期时间: {TokenExpiry}");
            return true;
        }

        /// <summary>
        /// 用户登录（同步版本，用于兼容现有代码）
        /// </summary>
        public bool Login(string userName, string password)
        {
            // 使用同步方式调用异步方法
            var task = Task.Run(() => LoginAsync(userName, password));
            return task.Result;
        }

        /// <summary>
        /// 用户注册
        /// </summary>
        public async Task<bool> RegisterAsync(string phone, string smsCode, string password)
        {
            var result = await RegisterWithMessageAsync(phone, smsCode, password).ConfigureAwait(false);
            return result.Success;
        }

        /// <summary>
        /// 用户注册（含错误文案，供登录窗 Bridge 使用）
        /// </summary>
        public async Task<RegisterResult> RegisterWithMessageAsync(string phone, string smsCode, string password)
        {
            try
            {
                return await ExecuteWithRetry(async () =>
                {
                    var registerRequest = new RegisterRequest
                    {
                        phone = phone,
                        sms_code = smsCode,
                        password = password
                    };

                    var json = JsonConvert.SerializeObject(registerRequest);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    var response = await _httpClient.PostAsync($"{BaseUrl}{RegisterEndpoint}", content);
                    var responseContent = await response.Content.ReadAsStringAsync();

                    System.Diagnostics.Debug.WriteLine($"注册API响应 - 状态码: {response.StatusCode}, 内容: {responseContent}");

                    if (!response.IsSuccessStatusCode)
                    {
                        var detail = TryExtractErrorDetail(responseContent);
                        return RegisterResult.Fail(detail ?? $"注册失败: HTTP {(int)response.StatusCode}");
                    }

                    var apiResponse = JsonConvert.DeserializeObject<ApiResponse<object>>(responseContent);

                    if (apiResponse != null)
                    {
                        if (apiResponse.success)
                        {
                            string message = apiResponse.message?.ToLower() ?? string.Empty;

                            if (message.Contains("失败") || message.Contains("错误") ||
                                message.Contains("fail") || message.Contains("error") ||
                                message.Contains("手机号已被注册") || message.Contains("用户名已存在") ||
                                message.Contains("username exists") ||
                                message.Contains("用户已注册") || message.Contains("user already registered"))
                            {
                                return RegisterResult.Fail(apiResponse.message ?? "注册失败");
                            }

                            return RegisterResult.Ok(apiResponse.message ?? "用户注册成功");
                        }

                        return RegisterResult.Fail(apiResponse.message ?? "注册失败");
                    }

                    return RegisterResult.Fail("注册失败，无法解析服务器响应");
                });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] RegisterWithMessageAsync 异常: {ex.Message}");
                return RegisterResult.Fail("注册失败，请稍后重试");
            }
        }

        private static string TryExtractErrorDetail(string responseContent)
        {
            if (string.IsNullOrWhiteSpace(responseContent))
            {
                return null;
            }

            try
            {
                var error = JsonConvert.DeserializeObject<ErrorResponse>(responseContent);
                if (!string.IsNullOrWhiteSpace(error?.detail))
                {
                    return error.detail;
                }
            }
            catch (JsonException)
            {
                // ignore
            }

            try
            {
                var api = JsonConvert.DeserializeObject<ApiResponse<object>>(responseContent);
                if (!string.IsNullOrWhiteSpace(api?.message))
                {
                    return api.message;
                }
            }
            catch (JsonException)
            {
                // ignore
            }

            return null;
        }

        /// <summary>
        /// 用户注册（同步版本已废弃旧签名；请用异步三参数注册）
        /// </summary>
        [Obsolete("Use RegisterAsync(phone, smsCode, password)")]
        public bool Register(string phone, string smsCode, string password)
        {
            var task = Task.Run(() => RegisterAsync(phone, smsCode, password));
            return task.Result;
        }

        /// <summary>
        /// 更新认证令牌信息（供 Auth 类使用）
        /// </summary>
        /// <param name="accessToken">新的访问令牌</param>
        /// <param name="refreshToken">新的刷新令牌</param>
        /// <param name="tokenExpiry">令牌过期时间</param>
        public void UpdateTokens(string accessToken, string refreshToken, DateTime? tokenExpiry)
        {
            AccessToken = accessToken;
            RefreshToken = refreshToken;
            TokenExpiry = tokenExpiry;
            SaveToSettings();
        }

        public void Logout()
        {
            lock (_stateLock)
            {
                if (_isLoggingOut)
                {
                    return;
                }

                if (!IsLoggedIn &&
                    string.IsNullOrEmpty(AccessToken) &&
                    string.IsNullOrEmpty(RefreshToken))
                {
                    return;
                }

                _isLoggingOut = true;
            }

            try
            {
                var oldUserName = UserName;

                IsLoggedIn = false;
                UserName = null;
                UserId = null;
                AccessToken = null;
                RefreshToken = null;
                LoginTime = null;
                TokenExpiry = null;
                ClearWorkspaceRootCache();

                SaveToSettings();

                OnUserLoggedOut?.Invoke(this, new UserEventArgs(oldUserName, null));
            }
            finally
            {
                lock (_stateLock)
                {
                    _isLoggingOut = false;
                }
            }
        }

        /// <summary>
        /// 检查是否已登录
        /// </summary>
        public bool CheckLoginStatus()
        {
            // 检查登录状态和token是否过期
            if (!IsLoggedIn || string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(AccessToken))
            {
                return false;
            }

            // 检查token是否过期
            if (TokenExpiry.HasValue && DateTime.Now > TokenExpiry.Value)
            {
                System.Diagnostics.Debug.WriteLine("Token已过期，自动登出");
                Logout();
                return false;
            }

            return true;
        }

        /// <summary>
        /// 获取用户信息摘要
        /// </summary>
        public string GetUserInfo()
        {
            if (IsLoggedIn)
            {
                return $"用户: {UserName} (登录时间: {LoginTime?.ToString("yyyy-MM-dd HH:mm:ss")})";
            }
            return "未登录";
        }

        /// <summary>
        /// 登录后或启动时已登录：拉取设置、写入默认、解析有效路径。
        /// </summary>
        public async Task InitializeWorkspaceSettingsAsync(IWin32Window owner = null)
        {
            if (!CheckLoginStatus())
            {
                return;
            }

            await LoadUserSettingsAsync().ConfigureAwait(false);
            await EnsureDefaultWorkspaceRootAsync().ConfigureAwait(false);
            EnsureEffectiveRootInteractive(owner);
        }

        public async Task<bool> LoadUserSettingsAsync()
        {
            if (!CheckLoginStatus())
            {
                return false;
            }

            try
            {
                var result = await BackendApiClient.GetUserSettingsAsync().ConfigureAwait(false);
                if (!result.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"[USER] 拉取用户设置失败: {result.StatusCode} {result.Body}");
                    return false;
                }

                var api = JsonConvert.DeserializeObject<ApiResponse<UserSettingsData>>(result.Body);
                if (api?.data != null)
                {
                    WorkspaceRootConfigured = NormalizeConfiguredRoot(api.data.workspace_root);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] LoadUserSettingsAsync 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>GET /users/me/llm-quota — 套餐名、已用/限额（付费用周期）、散装余额。</summary>
        public async Task<(bool ok, string planName, string planCode, double usedInspiration, double quotaInspiration, double looseInspiration, string planExpiresAt, string error)> GetLlmQuotaAsync()
        {
            if (!CheckLoginStatus())
            {
                return (false, "", "", 0, 0, 0, "", "请先登录");
            }

            try
            {
                var result = await BackendApiClient.GetLlmQuotaAsync().ConfigureAwait(false);
                if (!result.Success)
                {
                    return (false, "", "", 0, 0, 0, "", $"拉取额度失败: {result.StatusCode}");
                }

                var api = JsonConvert.DeserializeObject<ApiResponse<LlmQuotaData>>(result.Body);
                if (api?.data == null)
                {
                    return (false, "", "", 0, 0, 0, "", "额度响应无效");
                }

                var used = api.data.daily_used_inspiration;
                var quota = api.data.daily_quota_inspiration;
                if (string.Equals(api.data.quota_period, "subscription", StringComparison.OrdinalIgnoreCase))
                {
                    if (api.data.period_used_inspiration.HasValue)
                    {
                        used = api.data.period_used_inspiration.Value;
                    }
                    if (api.data.period_quota_inspiration.HasValue)
                    {
                        quota = api.data.period_quota_inspiration.Value;
                    }
                }

                return (
                    true,
                    api.data.plan_name ?? "",
                    api.data.plan_code ?? "",
                    used,
                    quota,
                    api.data.loose_inspiration,
                    api.data.plan_expires_at ?? "",
                    null
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] GetLlmQuotaAsync 异常: {ex.Message}");
                return (false, "", "", 0, 0, 0, "", ex.Message);
            }
        }

        /// <summary>设置页打开时重试 D40 默认写入。</summary>
        public async Task RetryEnsureDefaultWorkspaceRootAsync()
        {
            if (!CheckLoginStatus() || !string.IsNullOrWhiteSpace(WorkspaceRootConfigured))
            {
                return;
            }

            await EnsureDefaultWorkspaceRootAsync().ConfigureAwait(false);
        }

        /// <summary>D40：workspace_root 为空时自动 PUT 默认路径至 MySQL。</summary>
        public async Task<bool> EnsureDefaultWorkspaceRootAsync()
        {
            if (!CheckLoginStatus())
            {
                return false;
            }

            if (!string.IsNullOrWhiteSpace(WorkspaceRootConfigured))
            {
                return true;
            }

            var saved = await SaveWorkspaceRootAsync(WorkspacePathResolver.DefaultRoot).ConfigureAwait(false);
            if (!saved)
            {
                // D40 PUT 失败不阻塞：内存暂存默认路径，设置页打开时重试
                WorkspaceRootConfigured = WorkspacePathResolver.DefaultRoot;
                System.Diagnostics.Debug.WriteLine("[USER] 自动写入默认工作目录失败，使用内存默认值");
            }

            return true;
        }

        public async Task<bool> SaveWorkspaceRootAsync(string path)
        {
            if (!CheckLoginStatus())
            {
                return false;
            }

            var normalized = (path ?? "").Trim().Replace('\\', '/');
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return false;
            }

            try
            {
                var result = await BackendApiClient.PutUserSettingsAsync(normalized).ConfigureAwait(false);
                if (!result.Success)
                {
                    System.Diagnostics.Debug.WriteLine($"[USER] 保存工作目录失败: {result.StatusCode} {result.Body}");
                    return false;
                }

                var api = JsonConvert.DeserializeObject<ApiResponse<UserSettingsData>>(result.Body);
                WorkspaceRootConfigured = api?.data?.workspace_root ?? normalized;
                EnsureEffectiveRootInteractive(null);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] SaveWorkspaceRootAsync 异常: {ex.Message}");
                return false;
            }
        }

        /// <summary>解析本机有效路径；无可用盘符时弹窗选目录（可传 owner 窗口）。</summary>
        public bool EnsureEffectiveRootInteractive(IWin32Window owner)
        {
            try
            {
                var configured = string.IsNullOrWhiteSpace(WorkspaceRootConfigured)
                    ? WorkspacePathResolver.DefaultRoot
                    : WorkspaceRootConfigured;
                WorkspaceRootEffective = WorkspacePathResolver.EnsureEffectiveRoot(configured);
                WorkspaceDirectorySelectionRequired = false;
                return true;
            }
            catch (NoWritableWorkspaceDriveException)
            {
                return PromptSelectWorkspaceDirectory(owner);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] 工作区路径解析失败: {ex.Message}");
                WorkspaceDirectorySelectionRequired = true;
                return false;
            }
        }

        public string GetWorkspaceRemapHintText()
        {
            if (string.IsNullOrWhiteSpace(WorkspaceRootEffective)
                || string.IsNullOrWhiteSpace(WorkspaceRootConfigured))
            {
                return null;
            }

            if (!WorkspacePathResolver.IsRemapped(WorkspaceRootConfigured, WorkspaceRootEffective))
            {
                return null;
            }

            return $"当前实际使用 {WorkspaceRootEffective.Replace('\\', '/')}";
        }

        private bool PromptSelectWorkspaceDirectory(IWin32Window owner)
        {
            using (var dialog = new FolderBrowserDialog())
            {
                dialog.Description = "本机无可用默认盘符，请选择工作目录";
                dialog.ShowNewFolderButton = true;

                var result = owner != null ? dialog.ShowDialog(owner) : dialog.ShowDialog();
                if (result != DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    WorkspaceDirectorySelectionRequired = true;
                    System.Diagnostics.Debug.WriteLine("[USER] 用户取消选择工作目录");
                    return false;
                }

                if (!WorkspacePathResolver.ValidateWorkspacePathForSave(dialog.SelectedPath, out var error))
                {
                    MessageBox.Show(error ?? "目录无效", "工作目录", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WorkspaceDirectorySelectionRequired = true;
                    return false;
                }

                var path = dialog.SelectedPath.Replace('\\', '/');
                var saved = Task.Run(() => SaveWorkspaceRootAsync(path)).GetAwaiter().GetResult();
                if (!saved)
                {
                    MessageBox.Show("保存工作目录失败，请稍后重试", "工作目录", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    WorkspaceDirectorySelectionRequired = true;
                    return false;
                }

                WorkspaceDirectorySelectionRequired = false;
                return true;
            }
        }

        private static string NormalizeConfiguredRoot(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return value.Trim().Replace('\\', '/');
        }

        private void ClearWorkspaceRootCache()
        {
            WorkspaceRootConfigured = null;
            WorkspaceRootEffective = null;
            WorkspaceDirectorySelectionRequired = false;
        }

        // 事件定义
        public event EventHandler<UserEventArgs> OnUserLoggedIn;
        public event EventHandler<UserEventArgs> OnUserLoggedOut;

        #region 私有方法

        private void LoadFromSettings()
        {
            try
            {
                IsLoggedIn = Properties.Settings.Default.IsLoggedIn;
                UserName = Properties.Settings.Default.UserName;
                AccessToken = Properties.Settings.Default.AccessToken;
                RefreshToken = Properties.Settings.Default.RefreshToken;

                // 解析Token过期时间
                if (!string.IsNullOrEmpty(Properties.Settings.Default.TokenExpiry))
                {
                    if (DateTime.TryParse(Properties.Settings.Default.TokenExpiry, out DateTime expiry))
                    {
                        TokenExpiry = expiry;
                    }
                }

                System.Diagnostics.Debug.WriteLine("[USER] 从设置加载用户状态:");
                System.Diagnostics.Debug.WriteLine($"[USER] IsLoggedIn: {IsLoggedIn}");
                System.Diagnostics.Debug.WriteLine($"[USER] UserName: {UserName ?? "[null]"}");
                System.Diagnostics.Debug.WriteLine($"[USER] AccessToken: {(string.IsNullOrEmpty(AccessToken) ? "[null]" : "[已设置]")}");
                System.Diagnostics.Debug.WriteLine($"[USER] RefreshToken: {(string.IsNullOrEmpty(RefreshToken) ? "[null]" : "[已设置]")}");
                System.Diagnostics.Debug.WriteLine($"[USER] TokenExpiry: {TokenExpiry?.ToString("yyyy-MM-dd HH:mm:ss") ?? "[null]"}");

                // 如果设置中标记为已登录但没有用户名或token，则重置状态
                if (IsLoggedIn && (string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(AccessToken)))
                {
                    System.Diagnostics.Debug.WriteLine("[USER] 登录状态不完整，重置为未登录");
                    IsLoggedIn = false;
                    SaveToSettings();
                }

                System.Diagnostics.Debug.WriteLine($"[USER] 最终登录状态: {CheckLoginStatus()}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[USER] 加载用户设置失败: {ex.Message}");
                // 出错时重置为未登录状态
                IsLoggedIn = false;
                UserName = null;
                AccessToken = null;
                RefreshToken = null;
                TokenExpiry = null;
            }
        }

        private void SaveToSettings()
        {
            try
            {
                Properties.Settings.Default.IsLoggedIn = IsLoggedIn;
                Properties.Settings.Default.UserName = UserName ?? "";
                Properties.Settings.Default.AccessToken = AccessToken ?? "";
                Properties.Settings.Default.RefreshToken = RefreshToken ?? "";
                Properties.Settings.Default.TokenExpiry = TokenExpiry?.ToString("O") ?? ""; // ISO 8601格式
                Properties.Settings.Default.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"保存用户设置失败: {ex.Message}");
            }
        }

        private bool ValidateCredentials(string userName, string password)
        {
            // 模拟验证逻辑
            // 在实际应用中，这里应该调用服务器API
            return !string.IsNullOrWhiteSpace(userName) &&
                   !string.IsNullOrWhiteSpace(password) &&
                   userName.Length >= 3 &&
                   password.Length >= 6;
        }

        private bool CanRegister(string userName, string password)
        {
            // 模拟注册逻辑
            // 在实际应用中，这里应该调用服务器API
            return !string.IsNullOrWhiteSpace(userName) &&
                   !string.IsNullOrWhiteSpace(password) &&
                   userName.Length >= 3 &&
                   password.Length >= 6;
        }

        private string GenerateUserId(string userName)
        {
            // 生成用户ID
            // 在实际应用中，用户ID应该从服务器获取
            return $"user_{Math.Abs(userName.GetHashCode())}";
        }

        /// <summary>
        /// 执行API调用并实现重试逻辑
        /// </summary>
        private async Task<T> ExecuteWithRetry<T>(Func<Task<T>> operation)
        {
            for (int attempt = 1; attempt <= ApiMaxRetries; attempt++)
            {
                try
                {
                    return await operation();
                }
                catch (HttpRequestException ex) when (attempt < ApiMaxRetries)
                {
                    System.Diagnostics.Debug.WriteLine($"API调用失败，重试 {attempt}/{ApiMaxRetries}: {ex.Message}");
                    await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, attempt))); // 指数退避
                }
                catch (TaskCanceledException ex) when (attempt < ApiMaxRetries)
                {
                    System.Diagnostics.Debug.WriteLine($"API调用超时，重试 {attempt}/{ApiMaxRetries}: {ex.Message}");
                    await Task.Delay(TimeSpan.FromSeconds(Math.Pow(2, attempt))); // 指数退避
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"API调用异常: {ex.Message}");
                    throw; // 非网络相关异常，直接抛出
                }
            }

            // 所有重试都失败，返回默认值
            return default(T);
        }

        #endregion
    }

    /// <summary>
    /// 用户事件参数
    /// </summary>
    public class UserEventArgs : EventArgs
    {
        public string UserName { get; }
        public string UserId { get; }

        public UserEventArgs(string userName, string userId)
        {
            UserName = userName;
            UserId = userId;
        }
    }

    public struct RegisterResult
    {
        public bool Success { get; set; }
        public string Message { get; set; }

        public static RegisterResult Ok()
        {
            return new RegisterResult { Success = true, Message = null };
        }

        public static RegisterResult Ok(string message)
        {
            return new RegisterResult { Success = true, Message = message };
        }

        public static RegisterResult Fail(string message)
        {
            return new RegisterResult { Success = false, Message = message };
        }
    }
}
