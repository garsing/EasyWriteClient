using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace WordAddIn1
{
    /// <summary>
    /// Vue 用户设置独立窗口（非模态，与 KnowledgeBaseForm 相同 WebView2 初始化方式）。
    /// </summary>
    public partial class UserSettingsForm : Form
    {
        private static UserSettingsForm _activeInstance;

        private WebView2 webView2;
        private WebView2Bridge bridge;
        private bool isDragging;
        private Point dragStartPoint;
        private readonly int _cornerRadius;

        public static bool IsOpen =>
            _activeInstance != null && !_activeInstance.IsDisposed;

        public static void ActivateExisting()
        {
            if (!IsOpen)
            {
                return;
            }

            _activeInstance.Activate();
            _activeInstance.BringToFront();
        }

        public static Task ShowAsync(IWin32Window owner)
        {
            var tcs = new TaskCompletionSource<bool>();
            var invokeTarget = owner as Control;

            void ShowOnUiThread()
            {
                try
                {
                    if (IsOpen)
                    {
                        ActivateExisting();
                        tcs.TrySetResult(true);
                        return;
                    }

                    var form = new UserSettingsForm();
                    form.ApplyDefaultWindowSize(invokeTarget);
                    form.FormClosed += (sender, args) =>
                    {
                        if (!form.IsDisposed)
                        {
                            form.Dispose();
                        }
                    };
                    form.Show();
                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }

            if (invokeTarget != null && !invokeTarget.IsDisposed)
            {
                invokeTarget.BeginInvoke(new Action(ShowOnUiThread));
            }
            else
            {
                ShowOnUiThread();
            }

            return tcs.Task;
        }

        public UserSettingsForm()
        {
            float dpiScale = 1f;
            try
            {
                using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
                {
                    dpiScale = g.DpiX / 96f;
                }
            }
            catch
            {
                // ignore
            }

            _cornerRadius = Math.Max(12, (int)Math.Round(14 * dpiScale));
            InitializeComponent();
            _activeInstance = this;
            FormClosed += UserSettingsForm_FormClosed;
            Shown += UserSettingsForm_Shown;
            SizeChanged += (_, __) => ApplyWindowRegion();
            ApplyWindowRegion();
        }

        private async void UserSettingsForm_Shown(object sender, EventArgs e)
        {
            Shown -= UserSettingsForm_Shown;
            ApplyWindowRegion();
            await InitializeWebView2Async();
        }

        /// <summary>
        /// 无边框窗体用 Region 做圆角，与桌面主窗一致。
        /// </summary>
        private void ApplyWindowRegion()
        {
            if (Width <= 0 || Height <= 0)
            {
                return;
            }

            Region old = Region;
            using (GraphicsPath path = CreateRoundedRectPath(new Rectangle(0, 0, Width, Height), _cornerRadius))
            {
                Region = new Region(path);
            }

            old?.Dispose();
        }

        private static GraphicsPath CreateRoundedRectPath(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            if (radius <= 0 || bounds.Width < d || bounds.Height < d)
            {
                path.AddRectangle(bounds);
                return path;
            }

            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void UserSettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (ReferenceEquals(_activeInstance, this))
            {
                _activeInstance = null;
            }
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = string.Empty;
            ApplyDefaultWindowSize();
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(880, 560);
            FormBorderStyle = FormBorderStyle.None;
            ShowIcon = false;
            BackColor = Color.FromArgb(236, 236, 236);

            webView2 = new WebView2
            {
                Dock = DockStyle.Fill,
                Name = "webView2",
                DefaultBackgroundColor = Color.FromArgb(236, 236, 236)
            };
            Controls.Add(webView2);

            MouseDown += UserSettingsForm_MouseDown;
            MouseMove += UserSettingsForm_MouseMove;
            MouseUp += UserSettingsForm_MouseUp;
            Name = "UserSettingsForm";
            ResumeLayout(false);
        }

        private void UserSettingsForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && e.Y < 56)
            {
                isDragging = true;
                dragStartPoint = new Point(e.X, e.Y);
            }
        }

        private void UserSettingsForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentScreenPos = PointToScreen(e.Location);
                Location = new Point(currentScreenPos.X - dragStartPoint.X, currentScreenPos.Y - dragStartPoint.Y);
            }
        }

        private void UserSettingsForm_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }

        /// <summary>
        /// 按当前显示器工作区比例设置默认窗口大小（宽约 62%，高约 68%）。
        /// </summary>
        private void ApplyDefaultWindowSize(Control owner = null)
        {
            Screen screen = null;
            if (owner != null && !owner.IsDisposed)
            {
                screen = Screen.FromControl(owner);
            }
            else
            {
                screen = Screen.PrimaryScreen;
            }

            Rectangle area = screen?.WorkingArea ?? new Rectangle(0, 0, 1920, 1080);
            int width = Math.Max(960, (int)Math.Round(area.Width * 0.62));
            int height = Math.Max(620, (int)Math.Round(area.Height * 0.68));
            Size = new Size(width, height);
        }

        private async Task InitializeWebView2Async()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAddIn1",
                    "WebView2DataSettings"
                );

                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
                await webView2.EnsureCoreWebView2Async(environment);

                string outputDir = AppDomain.CurrentDomain.BaseDirectory;
                string wwwrootPath = Path.Combine(outputDir, "wwwroot");
                string absoluteWwwrootPath = Path.GetFullPath(wwwrootPath);

                System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] wwwroot: {absoluteWwwrootPath}, 存在: {Directory.Exists(absoluteWwwrootPath)}");

                if (Directory.Exists(absoluteWwwrootPath))
                {
                    webView2.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        absoluteWwwrootPath,
                        CoreWebView2HostResourceAccessKind.Allow
                    );
                }

                webView2.CoreWebView2.Settings.AreDevToolsEnabled = true;
                webView2.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = true;
                webView2.CoreWebView2.Settings.IsScriptEnabled = true;
                webView2.CoreWebView2.Settings.IsWebMessageEnabled = true;

                bridge = new WebView2Bridge(webView2);
                RegisterMessageHandlers();

                webView2.CoreWebView2.NavigationCompleted += (navSender, navArgs) =>
                {
                    System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] 导航完成: {navArgs.IsSuccess}, URL: {webView2.CoreWebView2.Source}");
                };

                LoadHtmlContent();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] WebView2 初始化失败: {ex}");
                MessageBox.Show($"WebView2 初始化失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadHtmlContent()
        {
            try
            {
                if (webView2?.CoreWebView2 == null)
                {
                    return;
                }

                string outputDir = AppDomain.CurrentDomain.BaseDirectory;
                string htmlPath = Path.Combine(outputDir, "wwwroot", "settings.html");
                string absoluteHtmlPath = Path.GetFullPath(htmlPath);

                System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] 查找 HTML: {absoluteHtmlPath}, 存在: {File.Exists(absoluteHtmlPath)}");

                if (File.Exists(absoluteHtmlPath))
                {
                    string fileName = Path.GetFileName(absoluteHtmlPath);
                    string virtualUrl = $"http://appassets.local/{fileName}";
                    System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] 加载 HTML: {virtualUrl}");
                    webView2.CoreWebView2.Navigate(virtualUrl);
                }
                else
                {
                    LoadEmbeddedHtml();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[UserSettingsForm] 加载 HTML 失败: {ex.Message}");
                LoadEmbeddedHtml();
            }
        }

        private void LoadEmbeddedHtml()
        {
            if (webView2?.CoreWebView2 == null)
            {
                return;
            }

            string fallbackHtml = @"
<!DOCTYPE html>
<html lang=""zh-CN"">
<head>
    <meta charset=""UTF-8"">
    <title>用户设置</title>
    <style>
        body { margin: 0; padding: 20px; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; color: #666; }
    </style>
</head>
<body>
    <h2>用户设置</h2>
    <p>正在加载 Vue 应用...</p>
    <p>请确保已构建前端项目 (npm run build)</p>
</body>
</html>";

            webView2.CoreWebView2.NavigateToString(fallbackHtml);
        }

        private void RegisterMessageHandlers()
        {
            bridge.RegisterHandler("getUserSettings", async (data) =>
            {
                return await UserSettingsBridgeHandlers.GetUserSettingsAsync(this);
            });

            bridge.RegisterHandler("getLlmQuota", async (data) =>
            {
                return await UserSettingsBridgeHandlers.GetLlmQuotaAsync();
            });

            bridge.RegisterHandler("saveWorkspaceRoot", async (data) =>
            {
                return await UserSettingsBridgeHandlers.SaveWorkspaceRootAsync(data);
            });

            bridge.RegisterHandler("getInteractionSettings", async (data) =>
            {
                return await Task.FromResult(UserSettingsBridgeHandlers.GetInteractionSettings());
            });

            bridge.RegisterHandler("setInteractionSetting", async (data) =>
            {
                return await Task.FromResult(UserSettingsBridgeHandlers.SetInteractionSetting(data));
            });

            bridge.RegisterHandler("logoutUser", async (data) =>
            {
                return await UserSettingsBridgeHandlers.LogoutUserAsync(this);
            });

            bridge.RegisterHandler("closeSettingsWindow", async (data) =>
            {
                await UserSettingsBridgeHandlers.RunOnUiThreadAsync(this, () =>
                {
                    Close();
                });
                return new { success = true };
            });

            bridge.RegisterHandler("openLoginWindow", async (data) =>
            {
                return await HandleOpenLoginWindowAsync();
            });

            bridge.RegisterHandler("openExternalUrl", async (data) =>
            {
                return await UserSettingsBridgeHandlers.OpenExternalUrlAsync(data);
            });

            // 设置页直调 Backend（支付下单等）
            bridge.RegisterHandler("getApiConfig", async (data) =>
            {
                return KnowledgeBaseService.GetApiConfig(data);
            });
        }

        private Task<object> HandleOpenLoginWindowAsync()
        {
            var owner = FindForm();
            return LoginForm.ShowDialogAsync(owner, logoutFirst: true)
                .ContinueWith(_ => (object)new { success = true });
        }
    }
}
