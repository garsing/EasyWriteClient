using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using System.Windows.Forms;
using Newtonsoft.Json.Linq;

namespace WordAddIn1
{
    internal static class UserSettingsBridgeHandlers
    {
        public static async Task<object> GetUserSettingsAsync(Control uiOwner)
        {
            try
            {
                var userService = UserService.Instance;
                if (!userService.IsLoggedIn)
                {
                    return new { success = true, isLoggedIn = false };
                }

                await userService.LoadUserSettingsAsync().ConfigureAwait(false);
                await userService.EnsureDefaultWorkspaceRootAsync().ConfigureAwait(false);

                await RunOnUiThreadAsync(uiOwner, () =>
                {
                    userService.EnsureEffectiveRootInteractive(uiOwner?.FindForm());
                }).ConfigureAwait(false);

                return new
                {
                    success = true,
                    isLoggedIn = true,
                    username = userService.UserName ?? string.Empty,
                    workspaceRoot = userService.WorkspaceRootConfigured ?? WorkspacePathResolver.DefaultRoot,
                    remapHint = userService.GetWorkspaceRemapHintText() ?? string.Empty,
                    defaultRoot = WorkspacePathResolver.DefaultRoot
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsBridgeHandlers] 获取用户设置失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        public static async Task<object> GetLlmQuotaAsync()
        {
            try
            {
                var userService = UserService.Instance;
                if (!userService.IsLoggedIn)
                {
                    return new { success = false, message = "请先登录" };
                }

                var (ok, planName, planCode, usedInspiration, quotaInspiration, looseInspiration, planExpiresAt, error) =
                    await userService.GetLlmQuotaAsync().ConfigureAwait(false);
                if (!ok)
                {
                    return new { success = false, message = error ?? "加载额度失败" };
                }

                return new
                {
                    success = true,
                    planName = planName ?? string.Empty,
                    planCode = planCode ?? string.Empty,
                    planExpiresAt = planExpiresAt ?? string.Empty,
                    dailyUsedInspiration = usedInspiration,
                    dailyQuotaInspiration = quotaInspiration,
                    looseInspiration = looseInspiration
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsBridgeHandlers] 获取额度失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        public static async Task<object> SaveWorkspaceRootAsync(object data)
        {
            try
            {
                var userService = UserService.Instance;
                if (!userService.IsLoggedIn)
                {
                    return new { success = false, message = "请先登录" };
                }

                var path = ExtractStringField(data, "path");
                if (string.IsNullOrWhiteSpace(path))
                {
                    return new { success = false, message = "工作目录不能为空" };
                }

                if (!WorkspacePathResolver.ValidateWorkspacePathForSave(path, out var error))
                {
                    return new { success = false, message = error ?? "目录无效" };
                }

                var normalized = path.Trim().Replace('\\', '/');
                var ok = await userService.SaveWorkspaceRootAsync(normalized).ConfigureAwait(false);
                if (!ok)
                {
                    return new { success = false, message = "保存工作目录失败，请检查网络后重试" };
                }

                return new
                {
                    success = true,
                    workspaceRoot = userService.WorkspaceRootConfigured ?? normalized,
                    remapHint = userService.GetWorkspaceRemapHintText() ?? string.Empty
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsBridgeHandlers] 保存工作目录失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        public static async Task<object> LogoutUserAsync(Control uiOwner)
        {
            try
            {
                await RunOnUiThreadAsync(uiOwner, () =>
                {
                    UserService.Instance.Logout();
                }).ConfigureAwait(false);

                return new { success = true };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsBridgeHandlers] 退出登录失败: {ex.Message}");
                return new { success = false, message = ex.Message };
            }
        }

        public static string ExtractStringField(object data, string fieldName)
        {
            if (data is JObject jObj)
            {
                return jObj[fieldName]?.ToString();
            }

            if (data is Dictionary<string, object> dict && dict.ContainsKey(fieldName))
            {
                return dict[fieldName]?.ToString();
            }

            return null;
        }

        /// <summary>用系统默认浏览器打开外链（支付宝收银台等）。</summary>
        public static Task<object> OpenExternalUrlAsync(object data)
        {
            try
            {
                var url = ExtractStringField(data, "url")?.Trim();
                if (string.IsNullOrEmpty(url) ||
                    !(url.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                      url.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                {
                    return Task.FromResult<object>(new { success = false, message = "无效的 URL" });
                }

                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
                return Task.FromResult<object>(new { success = true });
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsBridgeHandlers] openExternalUrl 失败: {ex.Message}");
                return Task.FromResult<object>(new { success = false, message = ex.Message });
            }
        }

        public static Task RunOnUiThreadAsync(Control control, Action action)
        {
            var tcs = new TaskCompletionSource<bool>();

            void Run()
            {
                try
                {
                    action();
                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }

            if (control == null || control.IsDisposed)
            {
                Run();
                return tcs.Task;
            }

            if (control.InvokeRequired)
            {
                control.BeginInvoke(new Action(Run));
            }
            else
            {
                control.BeginInvoke(new Action(Run));
            }

            return tcs.Task;
        }
    }
}
