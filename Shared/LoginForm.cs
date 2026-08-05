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
    /// 必须在窗口 Shown（可见）后再 EnsureCoreWebView2 + Navigate，否则 Desktop 上常见整页白屏。
    /// </summary>
    public class LoginForm : Form
    {
        private static LoginForm _activeInstance;

        private WebView2 webView2;
        private WebView2Bridge bridge;
        private Button btnClose;
        private bool isDragging;
        private Point dragStartPoint;
        private bool _loadStarted;

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
            var tcs = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
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

                    if (logoutFirst)
                    {
                        UserService.Instance.Logout();
                    }

                    using (var form = new LoginForm())
                    {
                        // 不再在 ShowDialog 前初始化 WebView2（会导致白屏）
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
                invokeTarget.BeginInvoke(new Action(ShowOnUiThread));
            }
            else if (Application.OpenForms.Count > 0)
            {
                Application.OpenForms[0].BeginInvoke(new Action(ShowOnUiThread));
            }
            else
            {
                ShowOnUiThread();
            }

            return tcs.Task;
        }

        public LoginForm()
        {
            _activeInstance = this;
            FormClosed += LoginForm_FormClosed;
            Shown += LoginForm_Shown;
            InitializeComponent();
        }

        private void LoginForm_Shown(object sender, EventArgs e)
        {
            System.Diagnostics.Debug.WriteLine(
                $"[LoginForm] Shown ClientSize={ClientSize}, webView2={webView2?.Size}");
            // 可见后再初始化引擎并加载页面
            _ = InitAndLoadAsync();
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

            Text = "登录";
            // 适配常见笔记本分辨率；过大且在 Show 前 Init 的 WebView2 易白屏
            var screen = Screen.FromControl(this).WorkingArea;
            int w = Math.Min(920, Math.Max(720, screen.Width - 80));
            int h = Math.Min(720, Math.Max(560, screen.Height - 80));
            Size = new Size(w, h);
            MinimumSize = new Size(640, 480);
            StartPosition = FormStartPosition.CenterParent;
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
            btnClose.Click += (s, ev) => Close();
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

        private static void Log(string msg)
        {
            string line = DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg;
            System.Diagnostics.Debug.WriteLine("[LoginForm] " + line);
            try
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EasyWriteDesktop");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, "login-webview.log"),
                    line + Environment.NewLine);
            }
            catch
            {
            }
        }

        private async Task InitAndLoadAsync()
        {
            if (_loadStarted)
            {
                return;
            }

            _loadStarted = true;

            try
            {
                Log($"InitAndLoad Size={Size} ClientSize={ClientSize} webViewSize={webView2.Size}");

                // 强制布局，避免 Size=0 初始化
                PerformLayout();
                if (webView2.Width < 32 || webView2.Height < 32)
                {
                    webView2.Size = new Size(Math.Max(ClientSize.Width, 640), Math.Max(ClientSize.Height, 480));
                }

                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "EasyWriteDesktop",
                    "WebView2DataLogin");
                Directory.CreateDirectory(userDataFolder);

                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder)
                    .ConfigureAwait(true);
                await webView2.EnsureCoreWebView2Async(environment).ConfigureAwait(true);

                string wwwroot = Path.GetFullPath(
                    Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot"));
                string htmlPath = Path.Combine(wwwroot, "login.html");
                string jsPath = Path.Combine(wwwroot, "assets", "login.js");
                string cssPath = Path.Combine(wwwroot, "assets", "login.css");
                Log($"wwwroot={wwwroot} html={File.Exists(htmlPath)} js={File.Exists(jsPath)} css={File.Exists(cssPath)}");

                if (!Directory.Exists(wwwroot)
                    || !File.Exists(htmlPath)
                    || !File.Exists(jsPath)
                    || !File.Exists(cssPath))
                {
                    ShowFallbackHtml("未找到 login 前端文件，请在 Shared/Frontend 执行 npm run build 后重新生成 Desktop。");
                    return;
                }

                var core = webView2.CoreWebView2;
                core.Settings.AreDevToolsEnabled = true;
                core.Settings.AreDefaultScriptDialogsEnabled = true;
                core.Settings.IsScriptEnabled = true;
                core.Settings.IsWebMessageEnabled = true;
                webView2.ZoomFactor = 1.0;

                core.SetVirtualHostNameToFolderMapping(
                    "appassets.local",
                    wwwroot,
                    CoreWebView2HostResourceAccessKind.Allow);

                core.ProcessFailed += (_, ev) =>
                    Log("ProcessFailed kind=" + ev.ProcessFailedKind);

                var navTcs = new TaskCompletionSource<bool>(
                    TaskCreationOptions.RunContinuationsAsynchronously);
                core.NavigationCompleted += async (_, navArgs) =>
                {
                    Log("NavigationCompleted success=" + navArgs.IsSuccess
                        + " status=" + navArgs.WebErrorStatus
                        + " url=" + core.Source);
                    navTcs.TrySetResult(navArgs.IsSuccess);
                    if (!navArgs.IsSuccess)
                    {
                        ShowFallbackHtml("页面加载失败 (" + navArgs.WebErrorStatus + ")");
                        return;
                    }

                    try
                    {
                        await core.ExecuteScriptAsync(@"
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
})();").ConfigureAwait(true);
                    }
                    catch (Exception ex)
                    {
                        Log("fit script: " + ex.Message);
                    }
                };

                bridge = new WebView2Bridge(webView2);
                RegisterMessageHandlers();

                // 尺寸刷新后再导航（部分机器上首帧不绘制）
                var sz = webView2.Size;
                if (sz.Width > 0 && sz.Height > 0)
                {
                    webView2.Size = new Size(sz.Width + 1, sz.Height);
                    webView2.Size = sz;
                }

                Log("Navigate login.html");
                core.Navigate("http://appassets.local/login.html");
                btnClose.BringToFront();

                var finished = await Task.WhenAny(navTcs.Task, Task.Delay(12000)).ConfigureAwait(true);
                if (finished != navTcs.Task)
                {
                    Log("NavigationCompleted timeout");
                    ShowFallbackHtml("登录页加载超时。日志: %LOCALAPPDATA%\\EasyWriteDesktop\\login-webview.log");
                    return;
                }

                if (navTcs.Task.Result)
                {
                    await Task.Delay(500).ConfigureAwait(true);
                    try
                    {
                        string check = await core.ExecuteScriptAsync(
                            @"(function(){
  var app=document.getElementById('login-app');
  return JSON.stringify({
    href: location.href,
    appHtmlLen: app ? app.innerHTML.length : -1,
    readyState: document.readyState
  });
})();").ConfigureAwait(true);
                        Log("vue-check " + check);
                        if (check != null && (check.Contains("\"appHtmlLen\":0")
                            || check.Contains("正在加载登录界面")))
                        {
                            // 仍停在 loading 文案或空节点：打开 DevTools 便于排查
                            if (check.Contains("正在加载登录界面"))
                            {
                                Log("login Vue still on loading placeholder");
                                core.OpenDevToolsWindow();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("vue-check failed: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("InitAndLoad failed: " + ex);
                MessageBox.Show(
                    "登录界面初始化失败: " + ex.Message,
                    "错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
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
