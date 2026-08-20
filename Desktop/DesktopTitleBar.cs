using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace EasyWriteClient.Desktop
{
    /// <summary>
    /// WorkBuddy 风格顶栏：浅灰底、左图标+标题、右细线 Min/Max/Close；无菜单。
    /// </summary>
    internal sealed class DesktopTitleBar : Panel
    {
        private const int WmNclButtonDown = 0xA1;
        private const int HtCaption = 0x2;

        private readonly Label _titleLabel;
        private readonly TitleBarButton _btnLayout;
        private readonly TitleBarButton _btnMin;
        private readonly TitleBarButton _btnMax;
        private readonly TitleBarButton _btnClose;
        private readonly float _dpiScale;
        private readonly Image _logo;
        private readonly int _barHeightExpanded;
        private readonly int _barHeightCompact;
        private readonly int _btnWExpanded;
        private readonly int _btnWCompact;
        private readonly int _logoWidth;
        private MainForm _owner;

        public DesktopTitleBar(float dpiScale)
        {
            _dpiScale = dpiScale <= 0 ? 1f : dpiScale;
            _logo = LoadLogo();
            // 工作台约 Win11 高度；缩小版收窄
            _barHeightExpanded = Scale(32);
            _barHeightCompact = Scale(26);
            _btnWExpanded = Scale(36);
            _btnWCompact = Scale(30);
            int barHeight = _barHeightExpanded;
            int btnW = _btnWExpanded;
            int btnH = barHeight;

            Dock = DockStyle.Top;
            Height = barHeight;
            MinimumSize = new Size(0, barHeight);
            MaximumSize = new Size(0, barHeight);
            // 与侧栏同一 chrome 色，无底部分割线（WorkBuddy 一体框）
            BackColor = Color.FromArgb(0xF0, 0xF0, 0xF0);
            Padding = new Padding(Scale(8), 0, 0, 0);
            DoubleBuffered = true;

            // 仅 logo，不显示「易写」文字；区域用于拖拽（缩小版会隐藏）
            _logoWidth = Scale(28);
            if (_logo != null)
            {
                _logoWidth = Math.Max(
                    _logoWidth,
                    (int)Math.Round(Scale(18) * (_logo.Width / (double)_logo.Height)) + Scale(8));
            }

            _titleLabel = new Label
            {
                AutoSize = false,
                Text = string.Empty,
                BackColor = Color.Transparent,
                Dock = DockStyle.Left,
                Width = _logoWidth
            };
            _titleLabel.Paint += TitleLabel_Paint;
            _titleLabel.MouseDown += TitleBar_MouseDown;
            _titleLabel.MouseDoubleClick += TitleBar_MouseDoubleClick;

            _btnClose = new TitleBarButton(TitleBarButtonKind.Close, btnW, btnH, _dpiScale)
            {
                Dock = DockStyle.Right
            };
            _btnClose.Click += (_, __) => OwnerForm?.Close();

            _btnMax = new TitleBarButton(TitleBarButtonKind.Maximize, btnW, btnH, _dpiScale)
            {
                Dock = DockStyle.Right
            };
            _btnMax.Click += (_, __) =>
            {
                OwnerForm?.ToggleMaximizeRestore();
                SyncMaxButtonGlyph();
            };

            _btnMin = new TitleBarButton(TitleBarButtonKind.Minimize, btnW, btnH, _dpiScale)
            {
                Dock = DockStyle.Right
            };
            _btnMin.Click += (_, __) =>
            {
                if (OwnerForm == null)
                {
                    return;
                }

                // 缩小版：收起为圆形/贴边胶囊悬浮球；工作台：任务栏最小化
                if (OwnerForm.IsCompactLayout)
                {
                    OwnerForm.MinimizeToFloatBall();
                }
                else
                {
                    OwnerForm.WindowState = FormWindowState.Minimized;
                }
            };

            _btnLayout = new TitleBarButton(TitleBarButtonKind.LayoutToggle, btnW, btnH, _dpiScale)
            {
                Dock = DockStyle.Right
            };
            _btnLayout.Click += (_, __) =>
            {
                OwnerForm?.ToggleLayoutMode();
                SyncLayoutButton();
            };

            // Dock.Right：后添加的贴最右侧 → 视觉 [⇄][—][□][×]
            Controls.Add(_titleLabel);
            Controls.Add(_btnLayout);
            Controls.Add(_btnMin);
            Controls.Add(_btnMax);
            Controls.Add(_btnClose);

            MouseDown += TitleBar_MouseDown;
            MouseDoubleClick += TitleBar_MouseDoubleClick;
        }

        private MainForm OwnerForm => _owner ?? (_owner = FindForm() as MainForm);

        protected override void OnParentChanged(EventArgs e)
        {
            base.OnParentChanged(e);
            _owner = FindForm() as MainForm;
            ApplyChromeForLayout(_owner != null && _owner.IsCompactLayout);
        }

        /// <summary>
        /// 缩小版：隐藏 □（最大化）、保留 ⇄ / — / ×，收窄标题栏并隐藏左侧 Yi logo；工作台恢复全套按钮。
        /// 视觉顺序（Dock.Right 后添加更靠右）：[⇄][—][×]
        /// </summary>
        public void ApplyChromeForLayout(bool compact)
        {
            if (_btnMin != null)
            {
                _btnMin.Visible = true;
                _btnMin.ToolTipText = compact ? "最小化（悬浮球）" : "最小化";
            }

            if (_btnMax != null)
            {
                _btnMax.Visible = !compact;
            }

            if (_titleLabel != null)
            {
                // 缩小版去掉原生栏 Yi 图标
                _titleLabel.Visible = !compact;
                _titleLabel.Width = compact ? 0 : _logoWidth;
            }

            int barH = compact ? _barHeightCompact : _barHeightExpanded;
            int btnW = compact ? _btnWCompact : _btnWExpanded;
            Height = barH;
            MinimumSize = new Size(0, barH);
            MaximumSize = new Size(0, barH);
            ResizeTitleBarButton(_btnLayout, btnW, barH);
            ResizeTitleBarButton(_btnMin, btnW, barH);
            ResizeTitleBarButton(_btnMax, btnW, barH);
            ResizeTitleBarButton(_btnClose, btnW, barH);

            SyncLayoutButton();
            SyncMaxButtonGlyph();
            PerformLayout();
            Invalidate(true);
        }

        private static void ResizeTitleBarButton(TitleBarButton btn, int w, int h)
        {
            if (btn == null)
            {
                return;
            }

            btn.Width = w;
            btn.Height = h;
            btn.Invalidate();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _logo?.Dispose();
            }

            base.Dispose(disposing);
        }

        public void SyncMaxButtonGlyph()
        {
            if (_owner == null)
            {
                return;
            }

            _btnMax.Kind = _owner.IsCustomMaximized
                ? TitleBarButtonKind.Restore
                : TitleBarButtonKind.Maximize;
            _btnMax.Invalidate();
        }

        public void SyncLayoutButton()
        {
            if (_btnLayout == null)
            {
                return;
            }

            bool compact = _owner != null && _owner.IsCompactLayout;
            _btnLayout.ToolTipText = compact ? "展开窗口" : "缩小窗口";
            _btnLayout.Invalidate();
        }

        private void TitleBar_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            Form form = OwnerForm;
            if (form == null)
            {
                return;
            }

            // 自定义最大化时先还原再拖
            if (OwnerForm.IsCustomMaximized)
            {
                OwnerForm.ToggleMaximizeRestore();
                SyncMaxButtonGlyph();
            }

            ReleaseCapture();
            SendMessage(form.Handle, WmNclButtonDown, (IntPtr)HtCaption, IntPtr.Zero);
        }

        private void TitleBar_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            // 缩小版不最大化
            if (OwnerForm != null && OwnerForm.IsCompactLayout)
            {
                return;
            }

            OwnerForm?.ToggleMaximizeRestore();
            SyncMaxButtonGlyph();
        }

        private void TitleLabel_Paint(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.CompositingMode = CompositingMode.SourceOver;

            int h = Scale(18);
            int x = Scale(2);
            int y = Math.Max(0, (_titleLabel.Height - h) / 2);

            if (_logo != null)
            {
                int w = Math.Max(1, (int)Math.Round(h * (_logo.Width / (double)_logo.Height)));
                g.DrawImage(_logo, new Rectangle(x, y, w, h));
                return;
            }

            // 资源缺失时的占位
            using (var brush = new SolidBrush(Color.FromArgb(0x2D, 0x2D, 0x2D)))
            {
                g.FillEllipse(brush, x, y, h, h);
            }
        }

        private static Image LoadLogo() => LoadYiWriteLogo();

        /// <summary>与标题栏 / 悬浮球同源：yi-write_logo1.png。</summary>
        internal static Image LoadYiWriteLogo()
        {
            try
            {
                var asm = typeof(DesktopTitleBar).Assembly;
                string resName = asm.GetManifestResourceNames()
                    .FirstOrDefault(n => n.EndsWith("yi-write_logo1.png", StringComparison.OrdinalIgnoreCase));
                if (!string.IsNullOrEmpty(resName))
                {
                    using (Stream stream = asm.GetManifestResourceStream(resName))
                    {
                        if (stream != null)
                        {
                            using (Image img = Image.FromStream(stream))
                            {
                                return MakeWhiteTransparent(new Bitmap(img));
                            }
                        }
                    }
                }

                string besideExe = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Assets", "yi-write_logo1.png");
                if (File.Exists(besideExe))
                {
                    using (Image img = Image.FromFile(besideExe))
                    {
                        return MakeWhiteTransparent(new Bitmap(img));
                    }
                }
            }
            catch
            {
                // 忽略，回退到占位绘制
            }

            return null;
        }

        /// <summary>将近白背景像素设为透明，避免顶栏上出现白底方块。</summary>
        private static Bitmap MakeWhiteTransparent(Bitmap src)
        {
            for (int y = 0; y < src.Height; y++)
            {
                for (int x = 0; x < src.Width; x++)
                {
                    Color c = src.GetPixel(x, y);
                    if (c.A == 0)
                    {
                        continue;
                    }

                    if (c.R >= 240 && c.G >= 240 && c.B >= 240)
                    {
                        src.SetPixel(x, y, Color.FromArgb(0, c.R, c.G, c.B));
                    }
                }
            }

            return src;
        }

        private int Scale(int value)
        {
            return Math.Max(1, (int)Math.Round(value * _dpiScale));
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }

    internal enum TitleBarButtonKind
    {
        Minimize,
        Maximize,
        Restore,
        Close,
        LayoutToggle
    }

    internal sealed class TitleBarButton : Control
    {
        private bool _hover;
        private TitleBarButtonKind _kind;
        private readonly float _dpiScale;
        private string _toolTipText = string.Empty;
        private ToolTip _toolTip;

        public TitleBarButton(TitleBarButtonKind kind, int width, int height, float dpiScale)
        {
            _kind = kind;
            _dpiScale = dpiScale <= 0 ? 1f : dpiScale;
            Size = new Size(width, height);
            Cursor = Cursors.Hand;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);
            if (kind == TitleBarButtonKind.LayoutToggle)
            {
                _toolTip = new ToolTip { ShowAlways = true };
                ToolTipText = "缩小窗口";
            }
            else if (kind == TitleBarButtonKind.Minimize)
            {
                _toolTip = new ToolTip { ShowAlways = true };
                ToolTipText = "最小化";
            }
        }

        public string ToolTipText
        {
            get => _toolTipText;
            set
            {
                _toolTipText = value ?? string.Empty;
                if (_toolTip != null)
                {
                    _toolTip.SetToolTip(this, _toolTipText);
                }
            }
        }

        public TitleBarButtonKind Kind
        {
            get => _kind;
            set
            {
                if (_kind == value)
                {
                    return;
                }

                _kind = value;
                Invalidate();
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _hover = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _hover = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            Color bg = Color.FromArgb(0xF0, 0xF0, 0xF0);
            if (_hover)
            {
                bg = _kind == TitleBarButtonKind.Close
                    ? Color.FromArgb(0xE8, 0x11, 0x23)
                    : Color.FromArgb(0xE5, 0xE5, 0xE5);
            }

            using (var brush = new SolidBrush(bg))
            {
                g.FillRectangle(brush, ClientRectangle);
            }

            Color iconColor = _hover && _kind == TitleBarButtonKind.Close
                ? Color.White
                : Color.FromArgb(0x1F, 0x1F, 0x1F);

            using (var pen = new Pen(iconColor, Math.Max(1f, _dpiScale))
            {
                StartCap = LineCap.Flat,
                EndCap = LineCap.Flat
            })
            {
                float cx = Width / 2f;
                float cy = Height / 2f;
                float s = Math.Min(Width, Height) * 0.14f;

                switch (_kind)
                {
                    case TitleBarButtonKind.Minimize:
                        g.DrawLine(pen, cx - s, cy, cx + s, cy);
                        break;
                    case TitleBarButtonKind.Maximize:
                        g.DrawRectangle(pen, cx - s, cy - s, s * 2f, s * 2f);
                        break;
                    case TitleBarButtonKind.Restore:
                        float o = s * 0.35f;
                        g.DrawRectangle(pen, cx - s + o, cy - s - o, s * 1.7f, s * 1.7f);
                        using (var erase = new SolidBrush(bg))
                        {
                            g.FillRectangle(
                                erase,
                                cx - s - 0.5f,
                                cy - s + o - 0.5f,
                                s * 1.7f + 1f,
                                s * 1.7f + 1f);
                        }

                        g.DrawRectangle(pen, cx - s, cy - s + o, s * 1.7f, s * 1.7f);
                        break;
                    case TitleBarButtonKind.Close:
                        g.DrawLine(pen, cx - s, cy - s, cx + s, cy + s);
                        g.DrawLine(pen, cx + s, cy - s, cx - s, cy + s);
                        break;
                    case TitleBarButtonKind.LayoutToggle:
                        // ⇄：左右箭头（避免依赖字体缺字）
                        float aw = s * 1.6f;
                        float ah = s * 0.55f;
                        // 上箭头向右
                        g.DrawLine(pen, cx - aw, cy - ah, cx + aw * 0.35f, cy - ah);
                        g.DrawLine(pen, cx + aw * 0.05f, cy - ah - s * 0.45f, cx + aw * 0.35f, cy - ah);
                        g.DrawLine(pen, cx + aw * 0.05f, cy - ah + s * 0.45f, cx + aw * 0.35f, cy - ah);
                        // 下箭头向左
                        g.DrawLine(pen, cx + aw, cy + ah, cx - aw * 0.35f, cy + ah);
                        g.DrawLine(pen, cx - aw * 0.05f, cy + ah - s * 0.45f, cx - aw * 0.35f, cy + ah);
                        g.DrawLine(pen, cx - aw * 0.05f, cy + ah + s * 0.45f, cx - aw * 0.35f, cy + ah);
                        break;
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _toolTip?.Dispose();
                _toolTip = null;
            }

            base.Dispose(disposing);
        }
    }
}
