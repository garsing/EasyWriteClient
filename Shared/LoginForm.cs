using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace WordAddIn1
{
    /// <summary>
    /// Vue 登录/注册独立模态窗口（WebView2）。
    /// ShowDialog 前完成 WebView2 引擎初始化；Shown 后再 Navigate，避免模态阻塞导致白屏。
    /// </summary>
    public class LoginForm : Form
    {
        private static LoginForm _activeInstance;

        private WebView2 webView2;
        private WebView2Bridge bridge;
        private Button btnClose;
        private bool isDragging;
        private Point dragStartPoint;
        private bool _htmlLoaded;
        private TaskCompletionSource<bool> _initTcs = new TaskCompletionSource<bool>();

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

        public static Task ShowDialogAsync(IWin32Window owner, bool logoutFirst = false)
        {
            var tcs = new TaskCompletionSource<bool>();
            var invokeTarget = owner as Control;

            async void ShowOnUiThreadAsync()
            {
                try
                {
                    if (IsOpen)
                    {
                        ActivateExisting();
                        tcs.TrySetResult(true);
                        return;
                    }

                    if (logoutFirst)
                    {
                        UserService.Instance.Logout();
                    }

                    using (var form = new LoginForm())
                    {
                        System.Diagnostics.Debug.WriteLine("[LoginForm] 等待 WebView2 引擎就绪…");
                        await form.WaitForInitializationAsync().ConfigureAwait(true);
                        System.Diagnostics.Debug.WriteLine(
                            $"[LoginForm] 引擎就绪 ClientSize={form.ClientSize}，显示对话框");

                        if (owner != null)
                        {
                            form.ShowDialog(owner);
                        }
                        else
                        {
                            form.ShowDialog();
                        }
                    }

                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[LoginForm] ShowDialogAsync 失败: {ex}");
                    tcs.TrySetException(ex);
                }
            }

            if (invokeTarget != null && !invokeTarget.IsDisposed)
            {
                invokeTarget.BeginInvoke(new Action(ShowOnUiThreadAsync));
            }
            else if (Application.OpenForms.Count > 0)
            {
                Application.OpenForms[0].BeginInvoke(new Action(ShowOnUiThreadAsync));
            }
            else
            {
                ShowOnUiThreadAsync();
            }

            return tcs.Task;
        }

        public LoginForm()
        {
            _activeInstance = this;
            FormClosed += LoginForm_FormClosed;
            Shown += LoginForm_Shown;
            InitializeComponent();

            if (!IsHandleCreated)
            {
                CreateControl();
            }

            _ = PrepareWebViewEngineAsync();
        }

        public Task WaitForInitializationAsync()
        {
            return _initTcs.Task;
        }

        private void LoginForm_Shown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[LoginForm] Shown ClientSize={ClientSize}, webView2={webView2?.Size}");
            TryLoadHtmlContent();
        }

        private void LoginForm_FormClosed(object sender, FormClosedEventArgs e)
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
            Size = new Size(1080, 1040);
            MinimumSize = new Size(1080, 1040);
            MaximumSize = new Size(1080, 1040);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.None;
            ShowIcon = false;
            BackColor = Color.FromArgb(247, 247, 245);

            webView2 = new WebView2
            {
                Dock = DockStyle.Fill,
                Name = "webView2",
                DefaultBackgroundColor = Color.FromArgb(247, 247, 245)
            };
            Controls.Add(webView2);

            btnClose = new Button
            {
                Size = new Size(32, 32),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                FlatStyle = FlatStyle.Flat,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            btnClose.FlatAppearance.BorderSize = 0;
            btnClose.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 232, 228);
            btnClose.Click += (s, e) => Close();
            LoadCloseIcon();
            Controls.Add(btnClose);
            btnClose.BringToFront();

            MouseDown += LoginForm_MouseDown;
            MouseMove += LoginForm_MouseMove;
            MouseUp += LoginForm_MouseUp;
            Layout += LoginForm_Layout;

            Name = "LoginForm";
            ResumeLayout(false);
            LoginForm_Layout(this, null);
        }

        private void LoginForm_Layout(object sender, LayoutEventArgs e)
        {
            if (btnClose != null && ClientSize.Width > 0)
            {
                btnClose.Location = new Point(ClientSize.Width - btnClose.Width - 10, 12);
            }
        }

        private void LoadCloseIcon()
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                using (Stream closeStream = assembly.GetManifestResourceStream("close.png"))
                {
                    if (closeStream != null)
                    {
                        btnClose.BackgroundImage = Image.FromStream(closeStream);
                        btnClose.BackgroundImageLayout = ImageLayout.Zoom;
                    }
                    else
                    {
                        btnClose.Text = "×";
                        btnClose.Font = new Font("Segoe UI", 14, FontStyle.Regular);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginForm] 加载关闭图标失败: {ex.Message}");
                btnClose.Text = "×";
            }
        }

        /// <summary>
        /// 在 ShowDialog 之前完成 WebView2 引擎初始化（不 Navigate）。
        /// </summary>
        private async Task PrepareWebViewEngineAsync()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAddIn1",
                    "WebView2DataLogin"
                );

                System.Diagnostics.Debug.WriteLine($"[LoginForm] 创建 WebView2 环境…");
                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder)
                    .ConfigureAwait(true);
                await webView2.EnsureCoreWebView2Async(environment).ConfigureAwait(true);

                string outputDir = AppDomain.CurrentDomain.BaseDirectory;
                string wwwrootPath = Path.Combine(outputDir, "wwwroot");
                string absoluteWwwrootPath = Path.GetFullPath(wwwrootPath);

                System.Diagnostics.Debug.WriteLine($"[LoginForm] wwwroot: {absoluteWwwrootPath}");

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
                webView2.ZoomFactor = 1.0;

                webView2.CoreWebView2.NavigationCompleted += async (navSender, navArgs) =>
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[LoginForm] 导航完成: success={navArgs.IsSuccess}, url={webView2.CoreWebView2.Source}");
                    if (!navArgs.IsSuccess)
                    {
                        ShowFallbackHtml($"页面加载失败 ({navArgs.WebErrorStatus})");
                        return;
                    }

                    try
                    {
                        await webView2.CoreWebView2.ExecuteScriptAsync(@"
(function () {
  function fit() {
    var h = window.innerHeight, w = window.innerWidth;
    document.documentElement.style.cssText = 'width:'+w+'px;height:'+h+'px;overflow:hidden;margin:0;padding:0';
    document.body.style.cssText = 'width:'+w+'px;height:'+h+'px;overflow:hidden;margin:0;padding:0';
    var app = document.getElementById('login-app');
    if (app) app.style.cssText = 'width:'+w+'px;height:'+h+'px;overflow:hidden';
  }
  fit();
  window.addEventListener('resize', fit);
})();");
                    }
                    catch { }
                };

                bridge = new WebView2Bridge(webView2);
                RegisterMessageHandlers();

                _initTcs.TrySetResult(true);
                System.Diagnostics.Debug.WriteLine("[LoginForm] WebView2 引擎就绪");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[LoginForm] WebView2 初始化失败: {ex}");
                _initTcs.TrySetException(ex);
                MessageBox.Show($"WebView2 初始化失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TryLoadHtmlContent()
        {
            if (_htmlLoaded || webView2?.CoreWebView2 == null)
            {
                return;
            }

            _htmlLoaded = true;

            string outputDir = AppDomain.CurrentDomain.BaseDirectory;
            string htmlPath = Path.Combine(outputDir, "wwwroot", "login.html");
            string jsPath = Path.Combine(outputDir, "wwwroot", "assets", "login.js");
            string cssPath = Path.Combine(outputDir, "wwwroot", "assets", "login.css");

            System.Diagnostics.Debug.WriteLine(
                $"[LoginForm] login.html={File.Exists(htmlPath)}, login.js={File.Exists(jsPath)}, login.css={File.Exists(cssPath)}");

            if (!File.Exists(htmlPath) || !File.Exists(jsPath) || !File.Exists(cssPath))
            {
                ShowFallbackHtml("未找到 login 前端文件，请在 frontend 目录执行 npm run build 后重新生成。");
                return;
            }

            System.Diagnostics.Debug.WriteLine("[LoginForm] Navigate login.html");
            webView2.CoreWebView2.Navigate("http://appassets.local/login.html");
            btnClose.BringToFront();
        }

        private void ShowFallbackHtml(string message)
        {
            if (webView2?.CoreWebView2 == null)
            {
                return;
            }

            string safeMessage = System.Security.SecurityElement.Escape(message) ?? message;
            webView2.CoreWebView2.NavigateToString($@"
<!DOCTYPE html><html><head><meta charset=""UTF-8"">
<style>body{{margin:0;padding:40px;font-family:sans-serif;color:#666;background:#f7f7f5}}</style></head>
<body><h2>登录界面加载失败</h2><p>{safeMessage}</p></body></html>");
        }

        private void RegisterMessageHandlers()
        {
            bridge.RegisterHandler("loginUser", async (data) =>
            {
                return await LoginBridgeHandlers.LoginAsync(data, this);
            });

            bridge.RegisterHandler("loginBySms", async (data) =>
            {
                return await LoginBridgeHandlers.LoginBySmsAsync(data, this);
            });

            bridge.RegisterHandler("sendSmsCode", async (data) =>
            {
                return await LoginBridgeHandlers.SendSmsCodeAsync(data, this);
            });

            bridge.RegisterHandler("registerUser", async (data) =>
            {
                return await LoginBridgeHandlers.RegisterAsync(data, this);
            });

            bridge.RegisterHandler("closeLoginWindow", async (data) =>
            {
                return await LoginBridgeHandlers.CloseLoginWindowAsync(this);
            });

            bridge.RegisterHandler("openAgreement", async (data) =>
            {
                return await LoginBridgeHandlers.OpenAgreementAsync(data, this);
            });
        }

        private void LoginForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && e.Y < 56)
            {
                isDragging = true;
                dragStartPoint = new Point(e.X, e.Y);
            }
        }

        private void LoginForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Point currentScreenPos = PointToScreen(e.Location);
                Location = new Point(currentScreenPos.X - dragStartPoint.X, currentScreenPos.Y - dragStartPoint.Y);
            }
        }

        private void LoginForm_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }
    }
}
