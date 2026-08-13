using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// 缩小版闲置悬浮球：游离为正圆；贴左右边为横向半露胶囊（朝内一头为 Logo 圆）；悬停展开为圆；点击回主窗。
    /// </summary>
    internal sealed class FloatBallForm : Form
    {
        private enum DockSide
        {
            None,
            Left,
            Right
        }

        private readonly MainForm _owner;
        private readonly Image _logo;
        private readonly float _dpiScale;
        private readonly int _ballDiameter;
        private readonly int _formSide;
        /// <summary>横向胶囊完整宽度（贴边时约一半在屏外）。</summary>
        private readonly int _capsuleWidth;
        private readonly int _capsuleHeight;
        private readonly int _snapDistance;
        private Point _dragMouseScreen;
        private Point _dragFormLocation;
        private bool _dragging;
        private bool _moved;
        private DateTime _suppressActivateUntilUtc = DateTime.MinValue;
        private DockSide _dock = DockSide.None;
        private bool _hoverExpanded;

        public FloatBallForm(MainForm owner, float dpiScale)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _dpiScale = dpiScale <= 0 ? 1f : dpiScale;
            // 悬停/游离球稍大；贴边胶囊保持小（与球解耦）
            _ballDiameter = Math.Max(44, (int)Math.Round(52 * _dpiScale));
            _capsuleHeight = Math.Max(18, (int)Math.Round(22 * _dpiScale));
            // 横胶囊：完整宽约 2×高；贴边中线贴缘 → 屏内约一半 + 圆头
            _capsuleWidth = Math.Max(_capsuleHeight + 12, (int)Math.Round(_capsuleHeight * 2.1));
            _snapDistance = Math.Max(24, (int)Math.Round(32 * _dpiScale));
            int minSide = Math.Max(
                SystemInformation.MinimumWindowSize.Width,
                SystemInformation.MinimumWindowSize.Height);
            _formSide = Math.Max(Math.Max(_ballDiameter, _capsuleWidth), minSide);
            _logo = DesktopTitleBar.LoadYiWriteLogo();

            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            // 任务栏由主窗（最小化）占用；球在前台时点底栏才能稳定还原
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            DoubleBuffered = true;
            Text = AppDisplayName.Value;
            ShowIcon = true;
            if (owner.Icon != null)
            {
                Icon = (Icon)owner.Icon.Clone();
            }

            MinimumSize = new Size(_formSide, _formSide);
            MaximumSize = new Size(_formSide, _formSide);
            ClientSize = new Size(_formSide, _formSide);

            ApplyHitRegion();

            MouseDown += OnMouseDown;
            MouseMove += OnMouseMove;
            MouseUp += OnMouseUp;
            MouseEnter += OnMouseEnter;
            MouseLeave += OnMouseLeave;
            Paint += OnPaint;
            Activated += OnActivated;
            FormClosing += OnFormClosing;
            Resize += OnResize;
        }

        private const int WmSysCommand = 0x0112;
        private const int ScMinimize = 0xF020;
        private const int ScRestore = 0xF120;
        private const int ScMaximize = 0xF030;

        public int FormSide => _formSide;

        private bool ShowAsCapsule =>
            _dock != DockSide.None && !_hoverExpanded && !_dragging;

        public void ShowAt(Point location)
        {
            // 展示瞬间会 Activated，短时忽略，避免一收球就立刻展回
            _suppressActivateUntilUtc = DateTime.UtcNow.AddMilliseconds(400);
            _hoverExpanded = false;
            ForceSquareSize();
            Location = location;
            InferDockFromLocation();
            if (_dock != DockSide.None)
            {
                SnapToDockEdge();
            }

            if (!Visible)
            {
                Show();
            }
            else
            {
                BringToFront();
            }

            ForceSquareSize();
            ApplyHitRegion();
            Invalidate();
        }

        public void HideBall()
        {
            if (Visible)
            {
                Hide();
            }

            _dock = DockSide.None;
            _hoverExpanded = false;
            _suppressActivateUntilUtc = DateTime.UtcNow.AddMilliseconds(400);
        }

        /// <summary>任务栏点击：球已前台时系统会发最小化，改为直接展回缩小版。</summary>
        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WmSysCommand)
            {
                int cmd = (int)(m.WParam.ToInt64() & 0xFFF0);
                if (cmd == ScMinimize || cmd == ScRestore || cmd == ScMaximize)
                {
                    RequestLeaveFloatBall();
                    return;
                }
            }

            base.WndProc(ref m);
        }

        protected override void SetBoundsCore(int x, int y, int width, int height, BoundsSpecified specified)
        {
            base.SetBoundsCore(x, y, _formSide, _formSide, specified);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _logo?.Dispose();
            }

            base.Dispose(disposing);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (ClientSize.Width != _formSide || ClientSize.Height != _formSide)
            {
                ForceSquareSize();
            }

            ApplyHitRegion();
        }

        private void ForceSquareSize()
        {
            if (ClientSize.Width != _formSide || ClientSize.Height != _formSide)
            {
                ClientSize = new Size(_formSide, _formSide);
            }
        }

        /// <summary>正圆在窗内的矩形（居中）。</summary>
        private Rectangle BallBounds
        {
            get
            {
                int ox = Math.Max(0, (_formSide - _ballDiameter) / 2);
                int oy = Math.Max(0, (_formSide - _ballDiameter) / 2);
                return new Rectangle(ox, oy, _ballDiameter, _ballDiameter);
            }
        }

        /// <summary>横向胶囊完整矩形（贴边时约一半在屏外）。</summary>
        private Rectangle CapsuleBounds
        {
            get
            {
                int y = Math.Max(0, (_formSide - _capsuleHeight) / 2);
                if (_dock == DockSide.Left)
                {
                    // 贴左：胶囊靠窗左，Logo 圆头在右侧（朝向屏内）
                    return new Rectangle(0, y, _capsuleWidth, _capsuleHeight);
                }

                // 贴右：胶囊靠窗右，Logo 圆头在左侧（朝向屏内）
                return new Rectangle(_formSide - _capsuleWidth, y, _capsuleWidth, _capsuleHeight);
            }
        }

        /// <summary>胶囊朝内一端的 Logo 圆（直径=胶囊高）。</summary>
        private Rectangle CapsuleHeadBounds
        {
            get
            {
                Rectangle cap = CapsuleBounds;
                if (_dock == DockSide.Left)
                {
                    return new Rectangle(cap.Right - _capsuleHeight, cap.Y, _capsuleHeight, _capsuleHeight);
                }

                return new Rectangle(cap.X, cap.Y, _capsuleHeight, _capsuleHeight);
            }
        }

        private Rectangle VisualBounds => ShowAsCapsule ? CapsuleBounds : BallBounds;

        private void ApplyHitRegion()
        {
            if (_formSide <= 0)
            {
                return;
            }

            using (var path = new GraphicsPath())
            {
                if (ShowAsCapsule)
                {
                    AddCapsulePath(path, CapsuleBounds);
                }
                else
                {
                    path.AddEllipse(BallBounds);
                }

                Region old = Region;
                Region = new Region(path);
                old?.Dispose();
            }
        }

        private static void AddCapsulePath(GraphicsPath path, Rectangle r)
        {
            int radius = Math.Max(2, r.Height / 2);
            AddRoundedRect(path, r, radius);
        }

        private static void AddRoundedRect(GraphicsPath path, Rectangle r, int radius)
        {
            int d = radius * 2;
            if (d > r.Width)
            {
                d = r.Width;
            }

            if (d > r.Height)
            {
                d = r.Height;
            }

            path.AddArc(r.X, r.Y, d, d, 180, 90);
            path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.Clear(Color.Magenta);

            var fill = Color.FromArgb(0xF0, 0xF0, 0xF0);
            var stroke = Color.FromArgb(0xC8, 0xC8, 0xC8);
            var headFill = Color.FromArgb(0xFA, 0xFA, 0xFA);

            if (ShowAsCapsule)
            {
                Rectangle cap = CapsuleBounds;
                Rectangle head = CapsuleHeadBounds;

                using (var path = new GraphicsPath())
                {
                    AddCapsulePath(path, cap);
                    using (var brush = new SolidBrush(fill))
                    {
                        g.FillPath(brush, path);
                    }

                    using (var pen = new Pen(stroke))
                    {
                        g.DrawPath(pen, path);
                    }
                }

                // 头部圆圈 + Logo
                using (var brush = new SolidBrush(headFill))
                {
                    g.FillEllipse(brush, head);
                }

                using (var pen = new Pen(stroke))
                {
                    g.DrawEllipse(pen, head.X, head.Y, head.Width - 1, head.Height - 1);
                }

                int pad = Math.Max(3, head.Width / 7);
                var dest = new Rectangle(
                    head.X + pad,
                    head.Y + pad,
                    Math.Max(1, head.Width - pad * 2),
                    Math.Max(1, head.Height - pad * 2));
                DrawLogo(g, dest);
            }
            else
            {
                Rectangle ball = BallBounds;
                using (var brush = new SolidBrush(fill))
                {
                    g.FillEllipse(brush, ball);
                }

                using (var pen = new Pen(stroke))
                {
                    g.DrawEllipse(pen, ball.X, ball.Y, ball.Width - 1, ball.Height - 1);
                }

                int pad = Math.Max(6, _ballDiameter / 6);
                var dest = new Rectangle(
                    ball.X + pad,
                    ball.Y + pad,
                    Math.Max(1, ball.Width - pad * 2),
                    Math.Max(1, ball.Height - pad * 2));
                DrawLogo(g, dest);
            }
        }

        private void DrawLogo(Graphics g, Rectangle dest)
        {
            if (_logo != null)
            {
                g.DrawImage(_logo, dest);
            }
            else
            {
                using (var brush = new SolidBrush(Color.FromArgb(0x2D, 0x2D, 0x2D)))
                {
                    g.FillEllipse(brush, dest);
                }
            }
        }

        private void OnMouseEnter(object sender, EventArgs e)
        {
            if (_dock == DockSide.None || _dragging)
            {
                return;
            }

            if (!_hoverExpanded)
            {
                _hoverExpanded = true;
                SnapToDockEdge();
                ApplyHitRegion();
                Invalidate();
            }
        }

        private void OnMouseLeave(object sender, EventArgs e)
        {
            if (_dock == DockSide.None || _dragging)
            {
                return;
            }

            // Capture 拖拽时也会 Leave，忽略
            if (Capture)
            {
                return;
            }

            if (_hoverExpanded)
            {
                _hoverExpanded = false;
                SnapToDockEdge();
                ApplyHitRegion();
                Invalidate();
            }
        }

        private void OnMouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            _dragging = true;
            _moved = false;
            _dragMouseScreen = Control.MousePosition;
            _dragFormLocation = Location;
            Capture = true;

            // 拖动中始终显示完整球
            if (ShowAsCapsule || _hoverExpanded)
            {
                _hoverExpanded = true;
                ApplyHitRegion();
                Invalidate();
            }
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging)
            {
                return;
            }

            Point now = Control.MousePosition;
            int dx = now.X - _dragMouseScreen.X;
            int dy = now.Y - _dragMouseScreen.Y;
            if (!_moved && (Math.Abs(dx) > 3 || Math.Abs(dy) > 3))
            {
                _moved = true;
                // 开始拖离贴边：先按球态夹紧
                _dock = DockSide.None;
                _hoverExpanded = false;
                ApplyHitRegion();
                Invalidate();
            }

            if (_moved)
            {
                Location = ClampFormLocation(
                    new Point(_dragFormLocation.X + dx, _dragFormLocation.Y + dy),
                    forceBall: true);
            }
        }

        private void OnMouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            bool wasDrag = _moved;
            _dragging = false;
            Capture = false;

            if (wasDrag)
            {
                _moved = false;
                TrySnapDockAfterDrag();
                WindowLayoutStore.SaveBallLocation(Location);
                ApplyHitRegion();
                Invalidate();
                return;
            }

            _moved = false;
            _owner.LeaveFloatBall();
        }

        private void OnActivated(object sender, EventArgs e)
        {
            if (DateTime.UtcNow < _suppressActivateUntilUtc)
            {
                return;
            }

            // 任务栏激活（球尚未前台时）：直接回缩小版
            BeginInvoke(new Action(RequestLeaveFloatBall));
        }

        private void OnResize(object sender, EventArgs e)
        {
            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
                RequestLeaveFloatBall();
            }
        }

        private void RequestLeaveFloatBall()
        {
            if (IsDisposed || !Visible)
            {
                return;
            }

            // 拖拽过程中不因激活抖动展窗
            if (_dragging || _moved)
            {
                return;
            }

            _owner.LeaveFloatBall();
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                _owner.LeaveFloatBall();
            }
        }

        private void TrySnapDockAfterDrag()
        {
            Rectangle wa = Screen.FromPoint(
                new Point(Left + _formSide / 2, Top + _formSide / 2)).WorkingArea;
            Rectangle ball = BallBounds;
            int ballLeft = Left + ball.X;
            int ballRight = ballLeft + ball.Width;
            int distLeft = ballLeft - wa.Left;
            int distRight = wa.Right - ballRight;

            if (distLeft <= _snapDistance && distLeft <= distRight)
            {
                _dock = DockSide.Left;
                _hoverExpanded = false;
                SnapToDockEdge();
                return;
            }

            if (distRight <= _snapDistance)
            {
                _dock = DockSide.Right;
                _hoverExpanded = false;
                SnapToDockEdge();
                return;
            }

            _dock = DockSide.None;
            _hoverExpanded = false;
            Location = ClampFormLocation(Location, forceBall: true);
        }

        private void InferDockFromLocation()
        {
            Rectangle wa = Screen.FromPoint(
                new Point(Left + _formSide / 2, Top + _formSide / 2)).WorkingArea;
            // 用窗中心相对工作区判断贴边（胶囊半露时球心可能在缘上）
            int centerX = Left + _formSide / 2;
            if (centerX - wa.Left <= _snapDistance + _capsuleWidth / 2)
            {
                _dock = DockSide.Left;
            }
            else if (wa.Right - centerX <= _snapDistance + _capsuleWidth / 2)
            {
                _dock = DockSide.Right;
            }
            else
            {
                Rectangle ball = BallBounds;
                int ballLeft = Left + ball.X;
                int ballRight = ballLeft + ball.Width;
                if (ballLeft - wa.Left <= _snapDistance)
                {
                    _dock = DockSide.Left;
                }
                else if (wa.Right - ballRight <= _snapDistance)
                {
                    _dock = DockSide.Right;
                }
                else
                {
                    _dock = DockSide.None;
                }
            }
        }

        /// <summary>
        /// 贴边：胶囊态时横向胶囊中线贴工作区左右缘（约一半在屏外）；悬停球态则整圆贴缘。
        /// </summary>
        private void SnapToDockEdge()
        {
            if (_dock == DockSide.None)
            {
                return;
            }

            Rectangle wa = Screen.FromPoint(
                new Point(Left + _formSide / 2, Top + _formSide / 2)).WorkingArea;
            Rectangle visual = ShowAsCapsule ? CapsuleBounds : BallBounds;
            int y = Math.Max(
                wa.Top - visual.Y,
                Math.Min(Top, wa.Bottom - visual.Y - visual.Height));

            int x;
            if (ShowAsCapsule)
            {
                // 半露出：胶囊水平中线落在屏幕缘上
                if (_dock == DockSide.Left)
                {
                    x = wa.Left - visual.X - visual.Width / 2;
                }
                else
                {
                    x = wa.Right - visual.X - visual.Width / 2;
                }
            }
            else if (_dock == DockSide.Left)
            {
                x = wa.Left - visual.X;
            }
            else
            {
                x = wa.Right - visual.X - visual.Width;
            }

            Location = new Point(x, y);
        }

        public Point ClampFormLocation(Point formLocation)
        {
            return ClampFormLocation(formLocation, forceBall: false);
        }

        private Point ClampFormLocation(Point formLocation, bool forceBall)
        {
            bool asCapsule = !forceBall && ShowAsCapsule;
            Rectangle visual = asCapsule ? CapsuleBounds : BallBounds;
            Rectangle wa = Screen.FromPoint(
                new Point(formLocation.X + _formSide / 2, formLocation.Y + _formSide / 2)).WorkingArea;

            int minLeft;
            int maxLeft;
            if (asCapsule)
            {
                // 允许半个胶囊出屏（贴边迷你态）
                minLeft = wa.Left - visual.X - visual.Width / 2;
                maxLeft = wa.Right - visual.X - visual.Width / 2;
            }
            else
            {
                minLeft = wa.Left - visual.X;
                maxLeft = wa.Right - visual.X - visual.Width;
            }

            int minTop = wa.Top - visual.Y;
            int maxTop = wa.Bottom - visual.Y - visual.Height;

            int x = Math.Max(minLeft, Math.Min(formLocation.X, maxLeft));
            int y = Math.Max(minTop, Math.Min(formLocation.Y, maxTop));
            return new Point(x, y);
        }

        /// <summary>默认贴工作区右侧，横向半露胶囊。</summary>
        public Point DefaultLocationOnScreen(Screen screen)
        {
            _dock = DockSide.Right;
            _hoverExpanded = false;
            Rectangle wa = (screen ?? Screen.PrimaryScreen).WorkingArea;
            Rectangle visual = CapsuleBounds;
            int formLeft = wa.Right - visual.X - visual.Width / 2;
            int formTop = wa.Top + Math.Max(0, (wa.Height - visual.Height) / 2) - visual.Y;
            return new Point(formLeft, formTop);
        }
    }
}
