using System;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using WordAddIn1;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// 易写 Desktop 主窗：自定义顶栏 + WebView2（任务侧栏 + 对话，host=desktop）。
    /// </summary>
    public sealed class MainForm : Form
    {
        private const int WmNcHitTest = 0x84;
        private const int HtClient = 1;
        private const int HtLeft = 10;
        private const int HtRight = 11;
        private const int HtTop = 12;
        private const int HtTopLeft = 13;
        private const int HtTopRight = 14;
        private const int HtBottom = 15;
        private const int HtBottomLeft = 16;
        private const int HtBottomRight = 17;

        private readonly DesktopChatSurface _chatSurface;
        private readonly DesktopTitleBar _titleBar;
        private readonly float _dpiScale;
        private readonly int _resizeBorder;
        private bool _started;
        private bool _customMaximized;
        private Rectangle _restoreBounds;

        public MainForm()
        {
            Text = AppDisplayName.Value;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            DoubleBuffered = true;

            _dpiScale = GetDpiScale();
            _resizeBorder = Math.Max(6, (int)Math.Round(6 * _dpiScale));
            MinimumSize = ScaleSize(960, 640, _dpiScale);
            Size = ScaleSize(1280, 800, _dpiScale);
            BackColor = Color.FromArgb(0xF0, 0xF0, 0xF0);
            _restoreBounds = Bounds;

            HostCallbacks.NotifyUserLoggedInAllAsync = async () =>
            {
                await Task.CompletedTask;
            };

            _titleBar = new DesktopTitleBar(_dpiScale);
            _chatSurface = new DesktopChatSurface();

            // 先 Fill 后 Top，保证顶栏占用上方区域
            Controls.Add(_chatSurface);
            Controls.Add(_titleBar);

            Shown += OnShown;
            FormClosed += (_, __) => WordHost.Shutdown();
        }

        internal bool IsCustomMaximized => _customMaximized;

        internal void ToggleMaximizeRestore()
        {
            if (_customMaximized)
            {
                Bounds = _restoreBounds;
                _customMaximized = false;
            }
            else
            {
                if (WindowState == FormWindowState.Minimized)
                {
                    WindowState = FormWindowState.Normal;
                }

                _restoreBounds = Bounds;
                Bounds = Screen.FromControl(this).WorkingArea;
                _customMaximized = true;
            }

            _titleBar.SyncMaxButtonGlyph();
            Invalidate();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmNcHitTest && !_customMaximized && WindowState == FormWindowState.Normal)
            {
                base.WndProc(ref m);
                if (m.Result == (IntPtr)HtClient)
                {
                    // 处理负坐标（多显示器）
                    int lp = m.LParam.ToInt32();
                    short x = (short)(lp & 0xFFFF);
                    short y = (short)((lp >> 16) & 0xFFFF);
                    Point p = PointToClient(new Point(x, y));
                    int b = _resizeBorder;
                    bool left = p.X <= b;
                    bool right = p.X >= ClientSize.Width - b;
                    bool top = p.Y <= b;
                    bool bottom = p.Y >= ClientSize.Height - b;

                    if (top && left)
                    {
                        m.Result = (IntPtr)HtTopLeft;
                        return;
                    }

                    if (top && right)
                    {
                        m.Result = (IntPtr)HtTopRight;
                        return;
                    }

                    if (bottom && left)
                    {
                        m.Result = (IntPtr)HtBottomLeft;
                        return;
                    }

                    if (bottom && right)
                    {
                        m.Result = (IntPtr)HtBottomRight;
                        return;
                    }

                    if (left)
                    {
                        m.Result = (IntPtr)HtLeft;
                        return;
                    }

                    if (right)
                    {
                        m.Result = (IntPtr)HtRight;
                        return;
                    }

                    if (top)
                    {
                        m.Result = (IntPtr)HtTop;
                        return;
                    }

                    if (bottom)
                    {
                        m.Result = (IntPtr)HtBottom;
                        return;
                    }
                }

                return;
            }

            base.WndProc(ref m);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (!_customMaximized)
            {
                using (var pen = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8)))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
                }
            }
        }

        private async void OnShown(object sender, EventArgs e)
        {
            if (_started)
            {
                return;
            }

            _started = true;
            _restoreBounds = Bounds;
            try
            {
                UseWaitCursor = true;
                // Dock：先布局顶栏，再让 Fill 的 WebView 落在其下方（勿 BringToFront 打乱 z-order）
                PerformLayout();
                await _chatSurface.InitializeAsync().ConfigureAwait(true);
                PerformLayout();
                UseWaitCursor = false;
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

        private static float GetDpiScale()
        {
            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                return g.DpiX / 96f;
            }
        }

        private static Size ScaleSize(int width, int height, float dpiScale)
        {
            return new Size(
                Math.Max(1, (int)Math.Round(width * dpiScale)),
                Math.Max(1, (int)Math.Round(height * dpiScale)));
        }
    }
}
