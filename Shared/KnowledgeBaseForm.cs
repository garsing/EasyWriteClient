using System;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Microsoft.Web.WebView2.Core;

namespace WordAddIn1
{
    /// <summary>
    /// Vue 知识库独立窗口（非模态；Plugin / Desktop 共用）。
    /// </summary>
    public partial class KnowledgeBaseForm : Form
    {
        private static KnowledgeBaseForm _activeInstance;

        private WebView2 webView2;
        private WebView2Bridge bridge;

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

                    var form = new KnowledgeBaseForm();
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

        public KnowledgeBaseForm()
        {
            _activeInstance = this;
            InitializeComponent();
            InitializeWebView2();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (ReferenceEquals(_activeInstance, this))
            {
                _activeInstance = null;
            }

            base.OnFormClosed(e);
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            this.Text = "知识库";
            this.Size = new Size(2800, 1500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MinimumSize = new Size(800, 600);
            this.ShowIcon = false;

            webView2 = new WebView2();
            webView2.Dock = DockStyle.Fill;
            webView2.Name = "webView2";
            this.Controls.Add(webView2);

            this.Name = "KnowledgeBaseForm";
            this.ResumeLayout(false);
        }

        private static string ResolveWwwroot()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string local = Path.GetFullPath(Path.Combine(baseDir, "wwwroot"));
            string sharedWww = Path.GetFullPath(Path.Combine(
                baseDir, "..", "..", "..", "..", "Shared", "Frontend", "wwwroot"));

            bool localOk = File.Exists(Path.Combine(local, "knowledge-base.html"));
            bool sharedOk = File.Exists(Path.Combine(sharedWww, "knowledge-base.html"));
            if (localOk && sharedOk)
            {
                DateTime localStamp = Stamp(local);
                DateTime sharedStamp = Stamp(sharedWww);
                return sharedStamp >= localStamp ? sharedWww : local;
            }

            return sharedOk ? sharedWww : local;
        }

        private static DateTime Stamp(string wwwroot)
        {
            string js = Path.Combine(wwwroot, "assets", "knowledge-base.js");
            if (File.Exists(js))
            {
                return File.GetLastWriteTimeUtc(js);
            }

            string html = Path.Combine(wwwroot, "knowledge-base.html");
            return File.Exists(html) ? File.GetLastWriteTimeUtc(html) : DateTime.MinValue;
        }

        private async void InitializeWebView2()
        {
            try
            {
                string userDataFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "WordAddIn1",
                    "WebView2Data"
                );

                var environment = await CoreWebView2Environment.CreateAsync(userDataFolder: userDataFolder);
                await webView2.EnsureCoreWebView2Async(environment);

                string wwwrootPath = ResolveWwwroot();
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] wwwroot: {wwwrootPath}, 存在: {Directory.Exists(wwwrootPath)}");

                if (Directory.Exists(wwwrootPath))
                {
                    webView2.CoreWebView2.SetVirtualHostNameToFolderMapping(
                        "appassets.local",
                        wwwrootPath,
                        CoreWebView2HostResourceAccessKind.Allow
                    );
                    System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] ✓ 已设置虚拟主机名映射: appassets.local -> {wwwrootPath}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] ✗ 错误: 未找到 wwwroot 目录: {wwwrootPath}");
                }

                webView2.CoreWebView2.Settings.AreDevToolsEnabled = true;
                webView2.CoreWebView2.Settings.AreDefaultScriptDialogsEnabled = true;
                webView2.CoreWebView2.Settings.IsScriptEnabled = true;
                webView2.CoreWebView2.Settings.IsWebMessageEnabled = true;
                WebView2VueHost.DisableUserZoom(webView2);

                bridge = new WebView2Bridge(webView2);
                RegisterMessageHandlers();

                webView2.CoreWebView2.NavigationCompleted += (sender, e) =>
                {
                    System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] 导航完成: {e.IsSuccess}, URL: {webView2.CoreWebView2.Source}");
                    if (!e.IsSuccess)
                    {
                        System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] 导航失败: {e.WebErrorStatus}");
                    }
                };

                LoadHtmlContent();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] WebView2 初始化失败: {ex.Message}");
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

                string wwwrootPath = ResolveWwwroot();
                string htmlPath = Path.Combine(wwwrootPath, "knowledge-base.html");
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] 查找 HTML: {htmlPath}, 存在: {File.Exists(htmlPath)}");

                if (File.Exists(htmlPath))
                {
                    try
                    {
                        webView2.CoreWebView2.SetVirtualHostNameToFolderMapping(
                            "appassets.local",
                            wwwrootPath,
                            CoreWebView2HostResourceAccessKind.Allow
                        );
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] ✗ 设置虚拟主机名映射失败: {ex.Message}");
                    }

                    string virtualUrl = "http://appassets.local/knowledge-base.html";
                    System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] 加载 HTML: {virtualUrl}");
                    webView2.CoreWebView2.Navigate(virtualUrl);
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("[KnowledgeBaseForm] 未找到构建后的 HTML 文件，使用内嵌 HTML");
                    LoadEmbeddedHtml();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[KnowledgeBaseForm] 加载 HTML 失败: {ex.Message}");
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
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>知识库</title>
    <style>
        body { margin: 0; padding: 0; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; }
        #knowledge-base-app { height: 100vh; }
    </style>
</head>
<body>
    <div id=""knowledge-base-app"">
        <div style=""padding: 20px; text-align: center; color: #999;"">
            <h1>知识库</h1>
            <p>正在加载 Vue 应用...</p>
            <p>请确保已构建前端项目 (npm run build)</p>
        </div>
    </div>
</body>
</html>";

            webView2.CoreWebView2.NavigateToString(fallbackHtml);
        }

        private void RegisterMessageHandlers()
        {
            bridge.RegisterHandler("getApiConfig", async (data) =>
            {
                return KnowledgeBaseService.GetApiConfig(data);
            });
        }
    }
}
