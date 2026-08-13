using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// 缩小版闲置后的圆形悬浮球（独立小窗）；可拖动；点击或任务栏激活 → 展回主窗。
    /// </summary>
    internal sealed class FloatBallForm : Form
    {
        private readonly MainForm _owner;
        private readonly Image _logo;
        private readonly float _dpiScale;
        private readonly int _ballDiameter;
        private readonly int _formSide;
        private Point _dragMouseScreen;
        private Point _dragFormLocation;
        private bool _dragging;
        private bool _moved;
        private bool _ignoreNextActivate = true;

        public FloatBallForm(MainForm owner, float dpiScale)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            _dpiScale = dpiScale <= 0 ? 1f : dpiScale;
            _ballDiameter = Math.Max(40, (int)Math.Round(52 * _dpiScale));
            int minSide = Math.Max(
                SystemInformation.MinimumWindowSize.Width,
                SystemInformation.MinimumWindowSize.Height);
            _formSide = Math.Max(_ballDiameter, minSide);
            _logo = DesktopTitleBar.LoadYiWriteLogo();

            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = true;
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

            ApplyCircleRegion();

            MouseDown += OnMouseDown;
            MouseMove += OnMouseMove;
            MouseUp += OnMouseUp;
            Paint += OnPaint;
            Activated += OnActivated;
            FormClosing += OnFormClosing;
        }

        public int FormSide => _formSide;

        public void ShowAt(Point location)
        {
            _ignoreNextActivate = true;
            ForceSquareSize();
            Location = location;
            if (!Visible)
            {
                Show();
            }
            else
            {
                BringToFront();
            }

            ForceSquareSize();
            ApplyCircleRegion();
        }

        public void HideBall()
        {
            if (Visible)
            {
                Hide();
            }
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

            ApplyCircleRegion();
        }

        private void ForceSquareSize()
        {
            if (ClientSize.Width != _formSide || ClientSize.Height != _formSide)
            {
                ClientSize = new Size(_formSide, _formSide);
            }
        }

        private Rectangle BallBounds
        {
            get
            {
                int ox = Math.Max(0, (_formSide - _ballDiameter) / 2);
                int oy = Math.Max(0, (_formSide - _ballDiameter) / 2);
                return new Rectangle(ox, oy, _ballDiameter, _ballDiameter);
            }
        }

        private void ApplyCircleRegion()
        {
            if (_formSide <= 0)
            {
                return;
            }

            Rectangle ball = BallBounds;
            using (var path = new GraphicsPath())
            {
                path.AddEllipse(ball);
                Region old = Region;
                Region = new Region(path);
                old?.Dispose();
            }
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.Clear(Color.Magenta);

            Rectangle ball = BallBounds;
            using (var brush = new SolidBrush(Color.FromArgb(0xF0, 0xF0, 0xF0)))
            {
                g.FillEllipse(brush, ball);
            }

            using (var pen = new Pen(Color.FromArgb(0xC8, 0xC8, 0xC8)))
            {
                g.DrawEllipse(pen, ball.X, ball.Y, ball.Width - 1, ball.Height - 1);
            }

            int pad = Math.Max(6, _ballDiameter / 6);
            var dest = new Rectangle(
                ball.X + pad,
                ball.Y + pad,
                Math.Max(1, ball.Width - pad * 2),
                Math.Max(1, ball.Height - pad * 2));
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
            }

            if (_moved)
            {
                Location = ClampFormLocation(
                    new Point(_dragFormLocation.X + dx, _dragFormLocation.Y + dy));
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
                Location = ClampFormLocation(Location);
                WindowLayoutStore.SaveBallLocation(Location);
                return;
            }

            _moved = false;
            _owner.LeaveFloatBall();
        }

        private void OnActivated(object sender, EventArgs e)
        {
            if (_ignoreNextActivate)
            {
                _ignoreNextActivate = false;
                return;
            }

            // 延迟判定：点球拖动时 Activated 会早于 MouseDown，不能立刻展球
            BeginInvoke(new Action(() =>
            {
                if (IsDisposed || !Visible)
                {
                    return;
                }

                if (_dragging || _moved || MouseButtons == MouseButtons.Left)
                {
                    return;
                }

                _owner.LeaveFloatBall();
            }));
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                _owner.LeaveFloatBall();
            }
        }

        /// <summary>
        /// 按「可见圆」贴边夹紧（窗体比圆大，若按整窗夹紧则圆到不了屏幕边缘）。
        /// </summary>
        public Point ClampFormLocation(Point formLocation)
        {
            Rectangle ball = BallBounds;
            Rectangle wa = Screen.FromPoint(
                new Point(formLocation.X + _formSide / 2, formLocation.Y + _formSide / 2)).WorkingArea;

            // 允许窗体超出工作区，只要可见圆仍在工作区内
            int minLeft = wa.Left - ball.X;
            int maxLeft = wa.Right - ball.X - ball.Width;
            int minTop = wa.Top - ball.Y;
            int maxTop = wa.Bottom - ball.Y - ball.Height;

            int x = Math.Max(minLeft, Math.Min(formLocation.X, maxLeft));
            int y = Math.Max(minTop, Math.Min(formLocation.Y, maxTop));
            return new Point(x, y);
        }

        /// <summary>默认：可见圆贴工作区右侧并垂直居中。</summary>
        public Point DefaultLocationOnScreen(Screen screen)
        {
            Rectangle wa = (screen ?? Screen.PrimaryScreen).WorkingArea;
            Rectangle ball = BallBounds;
            int margin = Math.Max(0, (int)Math.Round(2 * _dpiScale));
            int formLeft = wa.Right - ball.X - ball.Width - margin;
            int formTop = wa.Top + Math.Max(0, (wa.Height - ball.Height) / 2) - ball.Y;
            return ClampFormLocation(new Point(formLeft, formTop));
        }
    }
}
