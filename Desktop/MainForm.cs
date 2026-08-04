using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// 易写 Desktop 主窗：单一 WebView2（任务侧栏 + 对话，host=desktop）。
    /// </summary>
    public sealed class MainForm : Form
    {
        private readonly DesktopChatSurface _chatSurface;
        private bool _started;

        public MainForm()
        {
            Text = AppDisplayName.Value;
            StartPosition = FormStartPosition.CenterScreen;
            MinimumSize = new Size(960, 640);
            Size = new Size(1280, 800);
            BackColor = Color.White;

            HostCallbacks.NotifyUserLoggedInAllAsync = async () =>
            {
                // Desktop 单宿主：登录后由 UserService 事件驱动 WS 重连
                await Task.CompletedTask;
            };

            _chatSurface = new DesktopChatSurface();
            Controls.Add(_chatSurface);

            Shown += OnShown;
            FormClosed += (_, __) => WordHost.Shutdown();
        }

        private async void OnShown(object sender, EventArgs e)
        {
            if (_started)
            {
                return;
            }

            _started = true;
            try
            {
                UseWaitCursor = true;
                PerformLayout();
                _chatSurface.BringToFront();
                await _chatSurface.InitializeAsync().ConfigureAwait(true);
                UseWaitCursor = false;
                // 等主界面 WebView 就绪后再弹登录，避免模态对话框挡住首次绘制
                await _chatSurface.EnsureLoggedInAsync().ConfigureAwait(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "桌面界面初始化失败：\r\n" + ex.Message,
                    AppDisplayName.Value,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                UseWaitCursor = false;
            }
        }
    }
}
