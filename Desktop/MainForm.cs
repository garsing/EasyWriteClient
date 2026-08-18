using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
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
        private const int WmSetCursor = 0x20;
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
        private readonly int _cornerRadius;
        private bool _started;
        private bool _customMaximized;
        private Rectangle _restoreBounds;
        private Rectangle _expandedBounds;
        private bool _expandedWasCustomMaximized;
        private WindowLayoutMode _layoutMode = WindowLayoutMode.Expanded;
        private Action _requestCompactHandler;
        private Action _bringDesktopToFrontHandler;
        private Action _clearDesktopTopMostHandler;
        private readonly System.Windows.Forms.Timer _idleTimer;
        private DateTime _lastActivityUtc = DateTime.UtcNow;
        private bool _isFloatBall;
        private bool _compactUiBusy;
        private FloatBallForm _floatBall;
        private Rectangle _boundsBeforeFloat;
        private DateTime _lastActivityNoteUtc = DateTime.MinValue;

        public MainForm()
        {
            Text = AppDisplayName.Value;
            StartPosition = FormStartPosition.CenterScreen;
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            DoubleBuffered = true;
            ShowIcon = true;
            Icon = LoadAppIcon();
            SetStyle(ControlStyles.ResizeRedraw, true);

            _dpiScale = GetDpiScale();
            // 略加宽命中带；并用 Padding 留出不被 WebView 盖住的边缘（否则无双箭头）
            _resizeBorder = Math.Max(8, (int)Math.Round(8 * _dpiScale));
            // WorkBuddy 风格外窗圆角
            _cornerRadius = Math.Max(10, (int)Math.Round(12 * _dpiScale));
            MinimumSize = ScaleSize(960, 640, _dpiScale);
            Size = ScaleSize(1280, 800, _dpiScale);
            BackColor = Color.FromArgb(0xF0, 0xF0, 0xF0);
            // Dock.Fill 的 WebView 不会盖住 Padding，边缘命中落在 Form 上 → 系统显示 Size 光标
            Padding = new Padding(_resizeBorder);
            _restoreBounds = Bounds;
            _expandedBounds = Bounds;

            HostCallbacks.NotifyUserLoggedInAllAsync = async () =>
            {
                await Task.CompletedTask;
            };

            _requestCompactHandler = () =>
            {
                if (IsDisposed)
                {
                    return;
                }

                if (InvokeRequired)
                {
                    BeginInvoke(new Action(RequestCompact));
                    return;
                }

                RequestCompact();
            };
            HostCallbacks.RequestCompact = _requestCompactHandler;

            _bringDesktopToFrontHandler = () =>
            {
                if (IsDisposed)
                {
                    return;
                }

                if (InvokeRequired)
                {
                    BeginInvoke(_bringDesktopToFrontHandler);
                    return;
                }

                BringDesktopToFrontOverBrowser();
            };
            HostCallbacks.BringDesktopToFront = _bringDesktopToFrontHandler;

            _clearDesktopTopMostHandler = () =>
            {
                if (IsDisposed)
                {
                    return;
                }

                if (InvokeRequired)
                {
                    BeginInvoke(_clearDesktopTopMostHandler);
                    return;
                }

                TopMost = false;
            };
            HostCallbacks.ClearDesktopTopMost = _clearDesktopTopMostHandler;

            _titleBar = new DesktopTitleBar(_dpiScale);
            _chatSurface = new DesktopChatSurface();

            _idleTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            _idleTimer.Tick += IdleTimer_Tick;

            // 先 Fill 后 Top，保证顶栏占用上方区域
            Controls.Add(_chatSurface);
            Controls.Add(_titleBar);

            Shown += OnShown;
            FormClosing += (_, __) =>
            {
                if (_isFloatBall && _floatBall != null && _floatBall.Visible)
                {
                    WindowLayoutStore.SaveBallLocation(_floatBall.Location);
                }

                PersistLayoutState();
            };
            ResizeBegin += (_, __) => NoteUserActivity();
            ResizeEnd += (_, __) =>
            {
                NoteUserActivity();
                if (_layoutMode == WindowLayoutMode.Compact && !_isFloatBall)
                {
                    PersistLayoutState();
                }
            };
            Move += (_, __) => NoteUserActivity();
            MouseDown += (_, __) => NoteUserActivity();
            MouseMove += (_, __) => NoteUserActivityThrottled();
            KeyPreview = true;
            KeyDown += (_, __) => NoteUserActivity();
            _titleBar.MouseDown += (_, __) => NoteUserActivity();
            _chatSurface.MouseDown += (_, __) => NoteUserActivity();
            // 球态下主窗在任务栏为最小化：点底栏会还原 → 在此展回缩小版
            Resize += OnResizeWhileFloatBall;
            FormClosed += (_, __) =>
            {
                _idleTimer.Stop();
                _idleTimer.Dispose();
                if (_floatBall != null)
                {
                    _floatBall.Dispose();
                    _floatBall = null;
                }

                if (ReferenceEquals(HostCallbacks.RequestCompact, _requestCompactHandler))
                {
                    HostCallbacks.RequestCompact = null;
                }

                if (ReferenceEquals(HostCallbacks.BringDesktopToFront, _bringDesktopToFrontHandler))
                {
                    HostCallbacks.BringDesktopToFront = null;
                }

                if (ReferenceEquals(HostCallbacks.ClearDesktopTopMost, _clearDesktopTopMostHandler))
                {
                    HostCallbacks.ClearDesktopTopMost = null;
                }

                WordHost.Shutdown();
                ExcelHost.Shutdown();
            };
            ApplyWindowRegion();
        }

        internal bool IsCustomMaximized => _customMaximized;

        internal bool IsCompactLayout => _layoutMode == WindowLayoutMode.Compact;

        internal bool IsFloatBall => _isFloatBall;

        internal void ToggleMaximizeRestore()
        {
            if (_layoutMode == WindowLayoutMode.Compact)
            {
                // 缩小版不走自定义最大化，避免与贴右窄窗冲突
                return;
            }

            if (_customMaximized)
            {
                Bounds = _restoreBounds;
                _customMaximized = false;
                Padding = new Padding(_resizeBorder);
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
                Padding = Padding.Empty;
            }

            _titleBar.SyncMaxButtonGlyph();
            ApplyWindowRegion();
            Invalidate();
        }

        internal void ToggleLayoutMode()
        {
            SetLayoutMode(
                _layoutMode == WindowLayoutMode.Compact
                    ? WindowLayoutMode.Expanded
                    : WindowLayoutMode.Compact);
        }

        internal void RequestCompact()
        {
            if (_isFloatBall)
            {
                LeaveFloatBall();
            }

            SetLayoutMode(WindowLayoutMode.Compact);
        }

        /// <summary>易写浏览窗可见后：缩小版 Desktop 置顶，浮在浏览窗之上。</summary>
        internal void BringDesktopToFrontOverBrowser()
        {
            if (_isFloatBall)
            {
                LeaveFloatBall();
            }

            if (_layoutMode != WindowLayoutMode.Compact)
            {
                SetLayoutMode(WindowLayoutMode.Compact);
            }

            TopMost = true;
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }

            if (!Visible)
            {
                Show();
            }

            BringToFront();
            Activate();
            NoteUserActivity();
        }

        /// <summary>Agent 流式/请求忙碌：禁止收球；若已是球则弹回缩小版。</summary>
        internal void NotifyAgentBusy(bool busy)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => NotifyAgentBusy(busy)));
                return;
            }

            NoteUserActivity();
            if (busy && _isFloatBall)
            {
                LeaveFloatBall();
            }
        }

        /// <summary>缩小版历史弹出等 UI 忙碌（前端 compactUiBusy）。</summary>
        internal void SetCompactUiBusy(bool busy)
        {
            if (IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(() => SetCompactUiBusy(busy)));
                return;
            }

            _compactUiBusy = busy;
            NoteUserActivity();
        }

        internal void NoteUserActivity()
        {
            _lastActivityUtc = DateTime.UtcNow;
        }

        private void NoteUserActivityThrottled()
        {
            DateTime now = DateTime.UtcNow;
            if ((now - _lastActivityNoteUtc).TotalMilliseconds < 400)
            {
                return;
            }

            _lastActivityNoteUtc = now;
            _lastActivityUtc = now;
        }

        private void IdleTimer_Tick(object sender, EventArgs e)
        {
            if (_layoutMode != WindowLayoutMode.Compact || _isFloatBall || IsDisposed)
            {
                return;
            }

            if (!WindowLayoutStore.GetAutoFloatEnabled())
            {
                return;
            }

            if (IsIdleBlocked())
            {
                NoteUserActivity();
                return;
            }

            if ((DateTime.UtcNow - _lastActivityUtc).TotalSeconds >= 5)
            {
                EnterFloatBall();
            }
        }

        private bool IsIdleBlocked()
        {
            if (_compactUiBusy)
            {
                return true;
            }

            if (_chatSurface != null && _chatSurface.IsProcessing)
            {
                return true;
            }

            // 鼠标仍在主窗上：不计闲置（移出后重新累计 5s）
            if (IsMouseOverMainWindow())
            {
                return true;
            }

            foreach (Form f in OwnedForms)
            {
                if (f != null && !f.IsDisposed && f.Visible)
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>光标是否落在本窗矩形内（含标题栏；WebView 内亦算）。</summary>
        private bool IsMouseOverMainWindow()
        {
            if (!Visible || WindowState == FormWindowState.Minimized)
            {
                return false;
            }

            return Bounds.Contains(Control.MousePosition);
        }

        private void StartIdleWatch()
        {
            if (_layoutMode != WindowLayoutMode.Compact || _isFloatBall)
            {
                _idleTimer.Enabled = false;
                return;
            }

            NoteUserActivity();
            _idleTimer.Enabled = true;
        }

        private void StopIdleWatch()
        {
            _idleTimer.Enabled = false;
        }

        private void EnterFloatBall()
        {
            if (_isFloatBall || _layoutMode != WindowLayoutMode.Compact || IsDisposed)
            {
                return;
            }

            if (IsIdleBlocked())
            {
                return;
            }

            _boundsBeforeFloat = Bounds;
            _isFloatBall = true;
            StopIdleWatch();

            EnsureFloatBall();
            _floatBall.ShowAt(ResolveBallLocation());

            // 主窗最小化留在任务栏（球不占任务栏）。
            // 这样无论球是否前台，点底栏都是「还原主窗」，必能回到缩小版。
            ShowInTaskbar = true;
            if (!Visible)
            {
                Show();
            }

            WindowState = FormWindowState.Minimized;
        }

        internal void LeaveFloatBall()
        {
            if (!_isFloatBall || IsDisposed)
            {
                return;
            }

            if (InvokeRequired)
            {
                BeginInvoke(new Action(LeaveFloatBall));
                return;
            }

            // 先清标志，避免还原 Normal 时 Resize 重入
            _isFloatBall = false;

            if (_floatBall != null && _floatBall.Visible)
            {
                WindowLayoutStore.SaveBallLocation(_floatBall.Location);
                _floatBall.HideBall();
            }

            if (_layoutMode != WindowLayoutMode.Compact)
            {
                _layoutMode = WindowLayoutMode.Compact;
            }

            ShowInTaskbar = true;
            if (WindowState != FormWindowState.Normal)
            {
                WindowState = FormWindowState.Normal;
            }

            if (_boundsBeforeFloat.Width >= MinimumSize.Width && _boundsBeforeFloat.Height >= MinimumSize.Height)
            {
                Bounds = _boundsBeforeFloat;
            }

            if (!Visible)
            {
                Show();
            }

            Activate();
            BringToFront();
            NoteUserActivity();
            StartIdleWatch();
            ApplyWindowRegion();
        }

        private void OnResizeWhileFloatBall(object sender, EventArgs e)
        {
            if (!_isFloatBall || IsDisposed)
            {
                return;
            }

            // 任务栏点击还原最小化主窗
            if (WindowState == FormWindowState.Normal)
            {
                LeaveFloatBall();
            }
        }

        private FloatBallForm EnsureFloatBall()
        {
            if (_floatBall == null || _floatBall.IsDisposed)
            {
                _floatBall = new FloatBallForm(this, _dpiScale);
            }

            return _floatBall;
        }

        private Point ResolveBallLocation()
        {
            FloatBallForm ball = EnsureFloatBall();
            Point? saved = WindowLayoutStore.LoadBallLocation();
            if (saved.HasValue)
            {
                return ball.ClampFormLocation(saved.Value);
            }

            return ball.DefaultLocationOnScreen(Screen.FromControl(this));
        }

        /// <summary>
        /// 缩小版内展开/收起侧栏时调整窗宽：向左增减，尽量不挤占聊天区宽度。
        /// delta&gt;0 展开侧栏（变宽），delta&lt;0 收起侧栏（变窄）。单位：逻辑像素（与前端 CSS 一致，再按 DPI Scale）。
        /// </summary>
        internal void AdjustCompactWidthForSidebar(int deltaCssPx)
        {
            if (IsDisposed || _layoutMode != WindowLayoutMode.Compact || deltaCssPx == 0)
            {
                return;
            }

            if (InvokeRequired)
            {
                // 必须同步完成，否则前端 Promise 先返回、侧栏已切换，会出现一帧错位闪动
                Invoke(new Action(() => AdjustCompactWidthForSidebar(deltaCssPx)));
                return;
            }

            int delta = (int)Math.Round(deltaCssPx * _dpiScale);
            Rectangle wa = Screen.FromControl(this).WorkingArea;
            int newW = Math.Max(MinimumSize.Width, Width + delta);
            int newLeft = Left - delta;

            if (delta > 0)
            {
                // 向左长：贴左屏边时改为只加宽（可能略向右伸出，再夹紧）
                if (newLeft < wa.Left)
                {
                    newLeft = wa.Left;
                    newW = Math.Min(Width + delta, wa.Right - newLeft);
                }
                else if (newLeft + newW > wa.Right)
                {
                    newW = wa.Right - newLeft;
                }
            }
            else
            {
                // 收窄：右缘尽量不动（左缘右移）
                if (newW < MinimumSize.Width)
                {
                    newW = MinimumSize.Width;
                    newLeft = Left + Width - newW;
                }

                if (newLeft + newW > wa.Right)
                {
                    newLeft = wa.Right - newW;
                }

                if (newLeft < wa.Left)
                {
                    newLeft = wa.Left;
                }
            }

            Bounds = new Rectangle(newLeft, Top, newW, Height);
            PersistLayoutState();
            ApplyWindowRegion();
        }

        internal void SetLayoutMode(WindowLayoutMode mode, bool persist = true, bool notifyWeb = true)
        {
            if (IsDisposed)
            {
                return;
            }

            if (_isFloatBall && mode == WindowLayoutMode.Expanded)
            {
                LeaveFloatBall();
            }

            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }

            bool already = _layoutMode == mode;
            if (mode == WindowLayoutMode.Compact)
            {
                if (!already)
                {
                    _expandedWasCustomMaximized = _customMaximized;
                    _expandedBounds = _customMaximized ? _restoreBounds : Bounds;
                    if (_customMaximized)
                    {
                        _customMaximized = false;
                    }

                    // 仅首次进入缩小版套默认/记忆尺寸；已在缩小版时不强制改 Bounds（可自由调节）
                    ApplyCompactBoundsInitial();
                }
                else
                {
                    // 已是缩小版：只保证最小尺寸，不重置用户拖过的大小
                    MinimumSize = ScaleSize(320, 280, _dpiScale);
                    MaximumSize = Size.Empty;
                }

                _layoutMode = WindowLayoutMode.Compact;
            }
            else
            {
                StopIdleWatch();
                _compactUiBusy = false;
                TopMost = false;
                if (_layoutMode == WindowLayoutMode.Compact && !_isFloatBall)
                {
                    // 离开缩小版前记下当前尺寸，便于下次回来
                    WindowLayoutStore.Save(WindowLayoutMode.Expanded, Bounds);
                }

                ApplyExpandedBounds();
                _layoutMode = WindowLayoutMode.Expanded;
            }

            _titleBar.ApplyChromeForLayout(_layoutMode == WindowLayoutMode.Compact);
            ApplyWindowRegion();
            Invalidate();

            if (_layoutMode == WindowLayoutMode.Compact && !_isFloatBall)
            {
                StartIdleWatch();
            }
            else
            {
                StopIdleWatch();
            }

            if (persist)
            {
                PersistLayoutState();
            }

            if (notifyWeb)
            {
                _chatSurface.NotifyLayoutModeChanged(
                    _layoutMode == WindowLayoutMode.Compact ? "compact" : "expanded");
            }
        }

        private void PersistLayoutState()
        {
            if (_layoutMode == WindowLayoutMode.Compact)
            {
                WindowLayoutStore.Save(WindowLayoutMode.Compact, Bounds);
            }
            else
            {
                WindowLayoutStore.Save(WindowLayoutMode.Expanded);
            }
        }

        /// <summary>进入缩小版时的初始 Bounds：优先上次记忆，否则贴右默认 520×720（可再自由拖改）。</summary>
        private void ApplyCompactBoundsInitial()
        {
            MinimumSize = ScaleSize(320, 280, _dpiScale);
            MaximumSize = Size.Empty;
            Padding = new Padding(_resizeBorder);
            Rectangle wa = Screen.FromControl(this).WorkingArea;
            Rectangle? saved = WindowLayoutStore.LoadCompactBounds();
            if (saved.HasValue)
            {
                Rectangle r = saved.Value;
                // 夹在当前屏工作区内，避免多屏记忆跑飞
                int w = Math.Max(MinimumSize.Width, Math.Min(r.Width, wa.Width));
                int h = Math.Max(MinimumSize.Height, Math.Min(r.Height, wa.Height));
                int left = Math.Max(wa.Left, Math.Min(r.X, wa.Right - w));
                int top = Math.Max(wa.Top, Math.Min(r.Y, wa.Bottom - h));
                Bounds = new Rectangle(left, top, w, h);
                return;
            }

            int width = ScaleSize(520, 280, _dpiScale).Width;
            int preferredH = ScaleSize(520, 720, _dpiScale).Height;
            int margin = Math.Max(16, (int)Math.Round(24 * _dpiScale));
            int maxH = Math.Max(MinimumSize.Height, wa.Height - margin * 2);
            int height = Math.Min(preferredH, maxH);
            int leftDef = wa.Right - width;
            int topDef = wa.Top + Math.Max(margin, (wa.Height - height) / 2);
            Bounds = new Rectangle(leftDef, topDef, width, height);
        }

        private void ApplyExpandedBounds()
        {
            MinimumSize = ScaleSize(960, 640, _dpiScale);
            Rectangle target = _expandedBounds;
            if (target.Width < MinimumSize.Width || target.Height < MinimumSize.Height)
            {
                Size def = ScaleSize(1280, 800, _dpiScale);
                Rectangle wa = Screen.FromControl(this).WorkingArea;
                target = new Rectangle(
                    wa.Left + Math.Max(0, (wa.Width - def.Width) / 2),
                    wa.Top + Math.Max(0, (wa.Height - def.Height) / 2),
                    def.Width,
                    def.Height);
            }

            if (_expandedWasCustomMaximized)
            {
                _restoreBounds = target;
                Bounds = Screen.FromControl(this).WorkingArea;
                _customMaximized = true;
                Padding = Padding.Empty;
            }
            else
            {
                Bounds = target;
                _customMaximized = false;
                _restoreBounds = Bounds;
                Padding = new Padding(_resizeBorder);
            }

            _expandedWasCustomMaximized = false;
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            ApplyWindowRegion();
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmNcHitTest && !_customMaximized && WindowState == FormWindowState.Normal)
            {
                // 直接按客户区坐标判边缘（不依赖 base 结果），避免被子控件吃掉后无法改大小
                int lp = m.LParam.ToInt32();
                short x = (short)(lp & 0xFFFF);
                short y = (short)((lp >> 16) & 0xFFFF);
                Point p = PointToClient(new Point(x, y));
                int hit = HitTestResize(p);
                if (hit != HtClient)
                {
                    m.Result = (IntPtr)hit;
                    return;
                }

                base.WndProc(ref m);
                return;
            }

            // 无边框窗：显式设置 Size 光标，避免只返回 HT* 却仍显示箭头
            if (m.Msg == WmSetCursor && !_customMaximized && WindowState == FormWindowState.Normal)
            {
                int hit = (int)((long)m.LParam & 0xFFFF);
                Cursor c = CursorForHit(hit);
                if (c != null)
                {
                    Cursor.Current = c;
                    m.Result = (IntPtr)1;
                    return;
                }
            }

            base.WndProc(ref m);
        }

        private int HitTestResize(Point clientPt)
        {
            int b = _resizeBorder;
            // ClientSize 含 Padding 外缘；命中带用窗体客户区绝对坐标
            bool left = clientPt.X >= 0 && clientPt.X <= b;
            bool right = clientPt.X >= ClientSize.Width - b && clientPt.X < ClientSize.Width;
            bool top = clientPt.Y >= 0 && clientPt.Y <= b;
            bool bottom = clientPt.Y >= ClientSize.Height - b && clientPt.Y < ClientSize.Height;

            if (top && left) return HtTopLeft;
            if (top && right) return HtTopRight;
            if (bottom && left) return HtBottomLeft;
            if (bottom && right) return HtBottomRight;
            if (left) return HtLeft;
            if (right) return HtRight;
            if (top) return HtTop;
            if (bottom) return HtBottom;
            return HtClient;
        }

        private static Cursor CursorForHit(int hit)
        {
            switch (hit)
            {
                case HtLeft:
                case HtRight:
                    return Cursors.SizeWE;
                case HtTop:
                case HtBottom:
                    return Cursors.SizeNS;
                case HtTopLeft:
                case HtBottomRight:
                    return Cursors.SizeNWSE;
                case HtTopRight:
                case HtBottomLeft:
                    return Cursors.SizeNESW;
                default:
                    return null;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (_customMaximized)
            {
                return;
            }

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (var pen = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8)))
            using (var path = CreateRoundedRectPath(
                new Rectangle(0, 0, ClientSize.Width - 1, ClientSize.Height - 1),
                _cornerRadius))
            {
                e.Graphics.DrawPath(pen, path);
            }
        }

        /// <summary>
        /// 普通窗口用圆角 Region；最大化时取消圆角以铺满工作区。
        /// </summary>
        private void ApplyWindowRegion()
        {
            if (Width <= 0 || Height <= 0)
            {
                return;
            }

            Region old = Region;
            if (_customMaximized)
            {
                Region = null;
                old?.Dispose();
                return;
            }

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

        private async void OnShown(object sender, EventArgs e)
        {
            if (_started)
            {
                return;
            }

            _started = true;
            _restoreBounds = Bounds;
            _expandedBounds = Bounds;
            try
            {
                UseWaitCursor = true;
                // Dock：先布局顶栏，再让 Fill 的 WebView 落在其下方（勿 BringToFront 打乱 z-order）
                PerformLayout();
                ApplyWindowRegion();
                await _chatSurface.InitializeAsync().ConfigureAwait(true);
                PerformLayout();
                ApplyWindowRegion();

                // 每次启动都是 expanded，不恢复上次缩小版
                _chatSurface.NotifyLayoutModeChanged("expanded");
                _titleBar.ApplyChromeForLayout(compact: false);

                // 形态恢复后再挂 Monitor：已开 Word 建渠可再触发自动缩小（I5）
                _chatSurface.StartOpenFilesMonitorIfNeeded();

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

        private static Icon LoadAppIcon()
        {
            try
            {
                // 与 ApplicationIcon 一致：优先读 exe 内嵌图标（任务栏/Alt-Tab）
                Icon fromExe = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (fromExe != null)
                {
                    return fromExe;
                }

                string icoPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "yi-write.ico");
                if (File.Exists(icoPath))
                {
                    return new Icon(icoPath);
                }
            }
            catch
            {
                // 忽略，回退系统默认
            }

            return SystemIcons.Application;
        }
    }
}
