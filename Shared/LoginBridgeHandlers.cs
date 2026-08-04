using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    internal static class LoginBridgeHandlers
    {
        public static async Task<object> LoginAsync(object data, Form loginForm)
        {
            try
            {
                var credential = UserSettingsBridgeHandlers.ExtractStringField(data, "credential")?.Trim();
                var password = UserSettingsBridgeHandlers.ExtractStringField(data, "password");

                if (string.IsNullOrEmpty(credential) || string.IsNullOrEmpty(password))
                {
                    return new { success = false, message = "请输入手机号和密码" };
                }

                var result = await UserService.Instance.LoginWithMessageAsync(credential, password).ConfigureAwait(false);
                if (!result.Success)
                {
                    return new { success = false, message = result.Message ?? "手机号或密码错误" };
                }

                await HostCallbacks.RaiseNotifyUserLoggedInAllAsync().ConfigureAwait(false);

                await UserSettingsBridgeHandlers.RunOnUiThreadAsync(loginForm, () =>
                {
                    _ = UserService.Instance.InitializeWorkspaceSettingsAsync(loginForm);
                    if (loginForm != null && !loginForm.IsDisposed)
                    {
                        loginForm.Close();
                    }
                }).ConfigureAwait(false);

                return new { success = true };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginBridgeHandlers] 登录失败: {ex.Message}");
                return new { success = false, message = "登录失败，请稍后重试" };
            }
        }

        public static async Task<object> LoginBySmsAsync(object data, Form loginForm)
        {
            try
            {
                var phone = UserSettingsBridgeHandlers.ExtractStringField(data, "phone")?.Trim();
                var smsCode = UserSettingsBridgeHandlers.ExtractStringField(data, "sms_code")?.Trim();

                if (string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(smsCode))
                {
                    return new { success = false, message = "请输入手机号和验证码" };
                }

                var result = await UserService.Instance.LoginBySmsWithMessageAsync(phone, smsCode).ConfigureAwait(false);
                if (!result.Success)
                {
                    return new { success = false, message = result.Message ?? "验证码错误或已过期" };
                }

                await HostCallbacks.RaiseNotifyUserLoggedInAllAsync().ConfigureAwait(false);

                await UserSettingsBridgeHandlers.RunOnUiThreadAsync(loginForm, () =>
                {
                    _ = UserService.Instance.InitializeWorkspaceSettingsAsync(loginForm);
                    if (loginForm != null && !loginForm.IsDisposed)
                    {
                        loginForm.Close();
                    }
                }).ConfigureAwait(false);

                return new { success = true };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginBridgeHandlers] 验证码登录失败: {ex.Message}");
                return new { success = false, message = "登录失败，请稍后重试" };
            }
        }

        public static async Task<object> SendSmsCodeAsync(object data, Form loginForm)
        {
            try
            {
                var phone = UserSettingsBridgeHandlers.ExtractStringField(data, "phone")?.Trim();
                var scene = UserSettingsBridgeHandlers.ExtractStringField(data, "scene")?.Trim();

                if (string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(scene))
                {
                    return new { success = false, message = "请输入手机号" };
                }

                var result = await UserService.Instance.SendSmsCodeAsync(phone, scene).ConfigureAwait(false);
                if (!result.Success)
                {
                    return new { success = false, message = result.Message ?? "验证码发送失败" };
                }

                return new { success = true, message = result.Message ?? "验证码已发送" };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginBridgeHandlers] 发码失败: {ex.Message}");
                return new { success = false, message = "验证码发送失败，请稍后重试" };
            }
        }

        public static async Task<object> RegisterAsync(object data, Form loginForm)
        {
            try
            {
                var phone = UserSettingsBridgeHandlers.ExtractStringField(data, "phone")?.Trim();
                var smsCode = UserSettingsBridgeHandlers.ExtractStringField(data, "sms_code")?.Trim();
                var password = UserSettingsBridgeHandlers.ExtractStringField(data, "password");

                if (string.IsNullOrEmpty(phone) || string.IsNullOrEmpty(smsCode) || string.IsNullOrEmpty(password))
                {
                    return new { success = false, message = "请填写手机号、验证码和密码" };
                }

                var result = await UserService.Instance.RegisterWithMessageAsync(phone, smsCode, password).ConfigureAwait(false);
                if (!result.Success)
                {
                    return new { success = false, message = result.Message ?? "注册失败" };
                }

                // 注册成功后直接登录并关窗
                var loginResult = await UserService.Instance.LoginWithMessageAsync(phone, password).ConfigureAwait(false);
                if (!loginResult.Success)
                {
                    return new
                    {
                        success = false,
                        message = loginResult.Message ?? "注册成功，但自动登录失败，请手动登录"
                    };
                }

                await HostCallbacks.RaiseNotifyUserLoggedInAllAsync().ConfigureAwait(false);

                await UserSettingsBridgeHandlers.RunOnUiThreadAsync(loginForm, () =>
                {
                    _ = UserService.Instance.InitializeWorkspaceSettingsAsync(loginForm);
                    if (loginForm != null && !loginForm.IsDisposed)
                    {
                        loginForm.Close();
                    }
                }).ConfigureAwait(false);

                return new { success = true };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginBridgeHandlers] 注册失败: {ex.Message}");
                return new { success = false, message = "注册失败，请稍后重试" };
            }
        }

        public static Task<object> CloseLoginWindowAsync(Form loginForm)
        {
            return UserSettingsBridgeHandlers.RunOnUiThreadAsync(loginForm, () =>
            {
                if (loginForm != null && !loginForm.IsDisposed)
                {
                    loginForm.Close();
                }
            }).ContinueWith(_ => (object)new { success = true });
        }

        public static async Task<object> OpenAgreementAsync(object data, Form loginForm)
        {
            try
            {
                var kind = UserSettingsBridgeHandlers.ExtractStringField(data, "kind")?.Trim();
                if (string.IsNullOrEmpty(kind))
                {
                    return new { success = false, message = "协议类型无效" };
                }

                await AgreementForm.ShowAsync(loginForm, kind).ConfigureAwait(false);
                return new { success = true };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginBridgeHandlers] 打开协议失败: {ex.Message}");
                return new { success = false, message = "协议页面打开失败" };
            }
        }

    }
}
