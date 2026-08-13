using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// 缩小版闲置后的圆形悬浮球（独立小窗）；点击或任务栏激活 → 展回主窗。
    /// </summary>
    internal sealed class FloatBallForm : Form
    {
        private readonly MainForm _owner;
        private readonly Image _logo;
        /// <summary>可见圆直径（逻辑 52×DPI）。</summary>
        private readonly int _ballDiameter;
        /// <summary>窗体边长（至少达到系统最小窗宽，且保持正方形，避免被拉成扁椭圆）。</summary>
        private readonly int _formSide;
        private Point _dragStart;
        private bool _dragging;
        private bool _moved;
        private bool _ignoreNextActivate = true;

        public FloatBallForm(MainForm owner, float dpiScale)
        {
            _owner = owner ?? throw new ArgumentNullException(nameof(owner));
            float scale = dpiScale <= 0 ? 1f : dpiScale;
            _ballDiameter = Math.Max(40, (int)Math.Round(52 * scale));
            // ShowInTaskbar 的顶层窗有系统最小宽度，若宽>高再按窗裁圆会变成扁椭圆
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
            // 透明键：圆外区域不显示（Region 负责命中；底色避开 logo 近黑）
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

        /// <summary>用于贴边计算的外接方边长。</summary>
        public int FormSide => _formSide;

        public void ShowAt(Point location)
        {
            _ignoreNextActivate = true;
            ForceSquareSize();
            Location = location;
            // 主窗已 Hide，勿 Show(owner)，避免从属窗随主窗隐藏
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
            // 阻止系统把宽度拉大、高度不变
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

            // 整窗先铺透明色，再画圆（圆外靠 TransparencyKey 隐去）
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
            _dragStart = e.Location;
        }

        private void OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging)
            {
                return;
            }

            int dx = e.X - _dragStart.X;
            int dy = e.Y - _dragStart.Y;
            if (!_moved && (Math.Abs(dx) > 3 || Math.Abs(dy) > 3))
            {
                _moved = true;
            }

            if (_moved)
            {
                Location = new Point(Left + dx, Top + dy);
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
            _moved = false;

            if (wasDrag)
            {
                ClampToWorkingArea();
                WindowLayoutStore.SaveBallLocation(Location);
                return;
            }

            _owner.LeaveFloatBall();
        }

        private void OnActivated(object sender, EventArgs e)
        {
            if (_ignoreNextActivate)
            {
                _ignoreNextActivate = false;
                return;
            }

            // 任务栏点击激活球窗 → 展球
            if (!_dragging)
            {
                _owner.LeaveFloatBall();
            }
        }

        private void OnFormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                _owner.LeaveFloatBall();
            }
        }

        private void ClampToWorkingArea()
        {
            Rectangle wa = Screen.FromControl(this).WorkingArea;
            int x = Math.Max(wa.Left, Math.Min(Left, wa.Right - Width));
            int y = Math.Max(wa.Top, Math.Min(Top, wa.Bottom - Height));
            Location = new Point(x, y);
        }
    }
}
