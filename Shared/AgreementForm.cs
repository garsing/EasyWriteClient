using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.WinForms;

namespace WordAddIn1
{
    /// <summary>
    /// 展示用户协议 / 隐私政策的独立窗口（WebView2 加载 wwwroot/legal/*.html）。
    /// </summary>
    public class AgreementForm : Form
    {
        private static AgreementForm _activeInstance;

        private string _documentKind;
        private WebView2 webView2;
        private Button btnClose;
        private bool isDragging;
        private Point dragStartPoint;

        public static bool IsOpen =>
            _activeInstance != null && !_activeInstance.IsDisposed;

        public AgreementForm(string documentKind)
        {
            _documentKind = NormalizeKind(documentKind);
            InitializeComponent();
            _activeInstance = this;
            FormClosed += (_, __) =>
            {
                if (ReferenceEquals(_activeInstance, this))
                {
                    _activeInstance = null;
                }
            };
            Shown += AgreementForm_Shown;
        }

        public static Task ShowAsync(Form owner, string documentKind)
        {
            var kind = NormalizeKind(documentKind);
            var tcs = new TaskCompletionSource<bool>();

            void ShowOnUi()
            {
                try
                {
                    if (IsOpen)
                    {
                        _activeInstance._documentKind = kind;
                        _activeInstance.CenterOnOwner(owner);
                        _activeInstance.Activate();
                        _activeInstance.BringToFront();
                        _ = _activeInstance.NavigateToKindAsync(kind);
                        tcs.TrySetResult(true);
                        return;
                    }

                    var form = new AgreementForm(kind);
                    form.CenterOnOwner(owner);
                    if (owner != null && !owner.IsDisposed)
                    {
                        form.Show(owner);
                    }
                    else
                    {
                        form.Show();
                    }

                    tcs.TrySetResult(true);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                }
            }

            if (owner != null && !owner.IsDisposed && owner.InvokeRequired)
            {
                owner.BeginInvoke(new Action(ShowOnUi));
            }
            else
            {
                ShowOnUi();
            }

            return tcs.Task;
        }

        private static string NormalizeKind(string kind)
        {
            var value = (kind ?? string.Empty).Trim().ToLowerInvariant();
            return value == "privacy" ? "privacy" : "user";
        }

        /// <summary>
        /// 与登录/注册窗口中心对齐（协议窗更大时向四周扩展）。
        /// </summary>
        private void CenterOnOwner(Form owner)
        {
            if (owner == null || owner.IsDisposed)
            {
                StartPosition = FormStartPosition.CenterScreen;
                return;
            }

            StartPosition = FormStartPosition.Manual;
            int x = owner.Left + (owner.Width - Width) / 2;
            int y = owner.Top + (owner.Height - Height) / 2;
            Location = new Point(x, y);
        }

        private async void AgreementForm_Shown(object sender, EventArgs e)
        {
            Shown -= AgreementForm_Shown;
            await InitializeWebView2Async().ConfigureAwait(true);
        }

        private void InitializeComponent()
        {
            SuspendLayout();

            Text = string.Empty;
            // 比登录窗（1080×1040）更大，阅读协议更舒适
            Size = new Size(1280, 1180);
            MinimumSize = new Size(1280, 1180);
            MaximumSize = new Size(1280, 1180);
            StartPosition = FormStartPosition.Manual;
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

            MouseDown += AgreementForm_MouseDown;
            MouseMove += AgreementForm_MouseMove;
            MouseUp += AgreementForm_MouseUp;
            Layout += AgreementForm_Layout;

            Name = "AgreementForm";
            ResumeLayout(false);
            AgreementForm_Layout(this, null);
        }

        private void AgreementForm_Layout(object sender, LayoutEventArgs e)
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
            catch
            {
                btnClose.Text = "×";
                btnClose.Font = new Font("Segoe UI", 14, FontStyle.Regular);
            }
        }

        private async Task InitializeWebView2Async()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAddIn1",
                    "WebView2DataAgreement"
                );

                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder)
                    .ConfigureAwait(true);
                await webView2.EnsureCoreWebView2Async(environment).ConfigureAwait(true);

                string wwwrootPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot"));
                if (Directory.Exists(wwwrootPath))
                {
                    webView2.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        wwwrootPath,
                        CoreWebView2HostResourceAccessKind.Allow
                    );
                }

                webView2.CoreWebView2.Settings.AreDevToolsEnabled = true;
                WebView2VueHost.DisableUserZoom(webView2);

                await NavigateToKindAsync(_documentKind).ConfigureAwait(true);
                btnClose.BringToFront();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AgreementForm] WebView2 初始化失败: {ex}");
                MessageBox.Show($"协议页面加载失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Task NavigateToKindAsync(string kind)
        {
            if (webView2?.CoreWebView2 == null)
            {
                return Task.CompletedTask;
            }

            string fileName = kind == "privacy" ? "privacy-policy.html" : "user-agreement.html";
            string htmlPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "wwwroot", "legal", fileName);
            if (!File.Exists(htmlPath))
            {
                string title = kind == "privacy" ? "隐私政策" : "用户协议";
                webView2.CoreWebView2.NavigateToString(
                    $"<!DOCTYPE html><html><head><meta charset=\"UTF-8\"></head>" +
                    $"<body style=\"font-family:Microsoft YaHei,sans-serif;padding:40px;color:#666;background:#f7f7f5\">" +
                    $"<h2>{title}</h2><p>未找到协议文件，请重新构建前端后重试。</p></body></html>");
                return Task.CompletedTask;
            }

            webView2.CoreWebView2.Navigate($"http://appassets.local/legal/{fileName}");
            return Task.CompletedTask;
        }

        private void AgreementForm_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left && e.Y < 56)
            {
                isDragging = true;
                dragStartPoint = new Point(e.X, e.Y);
            }
        }

        private void AgreementForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                Location = new Point(
                    Location.X + e.X - dragStartPoint.X,
                    Location.Y + e.Y - dragStartPoint.Y);
            }
        }

        private void AgreementForm_MouseUp(object sender, MouseEventArgs e)
        {
            isDragging = false;
        }
    }
}
