using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace WordAddIn1.BrowserHost
{
    internal static class YiWriteBrowserChromeTheme
    {
        public static readonly Color Frame = Color.FromArgb(0xDE, 0xE1, 0xE6);
        public static readonly Color Toolbar = Color.White;
        public static readonly Color TabSelected = Color.White;
        public static readonly Color TabHover = Color.FromArgb(0xE8, 0xEA, 0xED);
        public static readonly Color Omnibox = Color.FromArgb(0xF1, 0xF3, 0xF4);
        public static readonly Color OmniboxFocus = Color.White;
        public static readonly Color OmniboxBorder = Color.FromArgb(0xC4, 0xC7, 0xC5);
        public static readonly Color Text = Color.FromArgb(0x1F, 0x1F, 0x1F);
        public static readonly Color TextMuted = Color.FromArgb(0x44, 0x47, 0x46);
        public static readonly Color CloseHover = Color.FromArgb(0xE8, 0x11, 0x23);
        public static readonly Color CaptionHover = Color.FromArgb(0xE8, 0xEA, 0xED);
        public static readonly Color PlusHover = Color.FromArgb(0xE8, 0xEA, 0xED);
    }

    internal sealed class YiWriteBrowserTabStrip : Control
    {
        private const int WmNclButtonDown = 0xA1;
        private const int HtCaption = 0x2;

        private readonly float _dpi;
        private readonly Font _tabFont;
        private readonly YiWriteBrowserCaptionButton _btnMin;
        private readonly YiWriteBrowserCaptionButton _btnMax;
        private readonly YiWriteBrowserCaptionButton _btnClose;

        private IList<YiWriteBrowserPage> _pages = new List<YiWriteBrowserPage>();
        private YiWriteBrowserPage _front;
        private int _hotTab = -1;
        private bool _hotClose;
        private bool _hotPlus;
        private bool _captionAwaitDrag;
        private Point _captionDownScreen;

        public YiWriteBrowserTabStrip(float dpiScale)
        {
            _dpi = dpiScale <= 0 ? 1f : dpiScale;
            _tabFont = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);

            SetStyle(
                ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw
                | ControlStyles.Selectable,
                true);

            Dock = DockStyle.Top;
            BackColor = YiWriteBrowserChromeTheme.Frame;
            TabStop = false;

            int btnW = Scale(46);
            _btnClose = new YiWriteBrowserCaptionButton(YiWriteBrowserCaptionKind.Close, _dpi)
            {
                Dock = DockStyle.Right,
                Width = btnW
            };
            _btnClose.Click += (s, e) =>
            {
                EventHandler handler = CloseWindowRequested;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            };

            _btnMax = new YiWriteBrowserCaptionButton(YiWriteBrowserCaptionKind.Maximize, _dpi)
            {
                Dock = DockStyle.Right,
                Width = btnW
            };
            _btnMax.Click += (s, e) =>
            {
                EventHandler handler = MaximizeRequested;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            };

            _btnMin = new YiWriteBrowserCaptionButton(YiWriteBrowserCaptionKind.Minimize, _dpi)
            {
                Dock = DockStyle.Right,
                Width = btnW
            };
            _btnMin.Click += (s, e) =>
            {
                EventHandler handler = MinimizeRequested;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }
            };

            Controls.Add(_btnMin);
            Controls.Add(_btnMax);
            Controls.Add(_btnClose);
            Height = Scale(36);
        }

        public event EventHandler<string> TabActivateRequested;

        public event EventHandler<string> TabCloseRequested;

        public event EventHandler NewTabRequested;

        public event EventHandler MinimizeRequested;

        public event EventHandler MaximizeRequested;

        public event EventHandler CloseWindowRequested;

        public void BindPages(IList<YiWriteBrowserPage> pages, YiWriteBrowserPage front)
        {
            _pages = pages ?? new List<YiWriteBrowserPage>();
            _front = front;
            Invalidate();
        }

        public void SetMaximized(bool maximized)
        {
            _btnMax.Kind = maximized
                ? YiWriteBrowserCaptionKind.Restore
                : YiWriteBrowserCaptionKind.Maximize;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _tabFont.Dispose();
            }

            base.Dispose(disposing);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (_btnClose == null || _btnMax == null || _btnMin == null)
            {
                return;
            }

            int h = Height;
            _btnClose.Height = h;
            _btnMax.Height = h;
            _btnMin.Height = h;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;

            using (var brush = new SolidBrush(YiWriteBrowserChromeTheme.Frame))
            {
                g.FillRectangle(brush, ClientRectangle);
            }

            LayoutTabs(out List<Rectangle> tabs, out Rectangle plus);
            for (int i = 0; i < tabs.Count; i++)
            {
                DrawTab(g, i, tabs[i]);
            }

            DrawPlus(g, plus);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_captionAwaitDrag && e.Button == MouseButtons.Left)
            {
                int dx = Cursor.Position.X - _captionDownScreen.X;
                int dy = Cursor.Position.Y - _captionDownScreen.Y;
                if (dx * dx + dy * dy >= 16)
                {
                    _captionAwaitDrag = false;
                    Capture = false;
                    YiWriteBrowserForm dragForm = FindForm() as YiWriteBrowserForm;
                    if (dragForm != null && dragForm.IsCustomMaximized)
                    {
                        dragForm.ToggleMaximizeRestore();
                    }

                    BeginCaptionDrag();
                }

                return;
            }

            HitTest(e.Location, out int tab, out bool close, out bool plus);
            if (tab == _hotTab && close == _hotClose && plus == _hotPlus)
            {
                return;
            }

            _hotTab = tab;
            _hotClose = close;
            _hotPlus = plus;
            Invalidate();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            _hotTab = -1;
            _hotClose = false;
            _hotPlus = false;
            Invalidate();
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            base.OnMouseClick(e);
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            HitTest(e.Location, out int tab, out bool close, out bool plus);
            if (plus)
            {
                EventHandler handler = NewTabRequested;
                if (handler != null)
                {
                    handler(this, EventArgs.Empty);
                }

                return;
            }

            if (tab < 0 || tab >= _pages.Count)
            {
                return;
            }

            YiWriteBrowserPage page = _pages[tab];
            if (page == null)
            {
                return;
            }

            if (close)
            {
                EventHandler<string> closeHandler = TabCloseRequested;
                if (closeHandler != null)
                {
                    closeHandler(this, page.TabUuid);
                }

                return;
            }

            EventHandler<string> activate = TabActivateRequested;
            if (activate != null)
            {
                activate(this, page.TabUuid);
            }
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            HitTest(e.Location, out int tab, out bool close, out bool plus);
            if (tab >= 0 || close || plus)
            {
                return;
            }

            YiWriteBrowserForm form = FindForm() as YiWriteBrowserForm;
            if (form == null)
            {
                return;
            }

            if (form.IsCustomMaximized)
            {
                _captionAwaitDrag = true;
                _captionDownScreen = Cursor.Position;
                Capture = true;
                return;
            }

            BeginCaptionDrag();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _captionAwaitDrag = false;
            Capture = false;
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            HitTest(e.Location, out int tab, out bool close, out bool plus);
            if (tab >= 0 || close || plus)
            {
                return;
            }

            EventHandler handler = MaximizeRequested;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void BeginCaptionDrag()
        {
            YiWriteBrowserForm form = FindForm() as YiWriteBrowserForm;
            if (form == null)
            {
                return;
            }

            ReleaseCapture();
            SendMessage(form.Handle, WmNclButtonDown, (IntPtr)HtCaption, IntPtr.Zero);
        }

        private void DrawTab(Graphics g, int index, Rectangle bounds)
        {
            if (index < 0 || index >= _pages.Count || bounds.Width < 8)
            {
                return;
            }

            YiWriteBrowserPage page = _pages[index];
            bool selected = page != null && page == _front;
            bool hot = index == _hotTab;

            if (selected || hot)
            {
                Color fill = selected
                    ? YiWriteBrowserChromeTheme.TabSelected
                    : YiWriteBrowserChromeTheme.TabHover;
                int radius = Scale(8);
                using (GraphicsPath path = RoundedTopRect(bounds, radius))
                using (var brush = new SolidBrush(fill))
                {
                    g.FillPath(brush, path);
                }
            }

            int closeSize = Scale(16);
            int closePad = Scale(8);
            var close = new Rectangle(
                bounds.Right - closePad - closeSize,
                bounds.Y + (bounds.Height - closeSize) / 2,
                closeSize,
                closeSize);

            bool showClose = selected || hot;
            if (showClose)
            {
                bool closeHot = hot && _hotClose;
                if (closeHot)
                {
                    using (var brush = new SolidBrush(YiWriteBrowserChromeTheme.CloseHover))
                    {
                        g.FillEllipse(brush, close);
                    }
                }

                using (var pen = new Pen(closeHot ? Color.White : YiWriteBrowserChromeTheme.TextMuted, 1.2f * _dpi))
                {
                    int inset = Scale(4);
                    g.DrawLine(
                        pen,
                        close.Left + inset,
                        close.Top + inset,
                        close.Right - inset,
                        close.Bottom - inset);
                    g.DrawLine(
                        pen,
                        close.Right - inset,
                        close.Top + inset,
                        close.Left + inset,
                        close.Bottom - inset);
                }
            }

            var textRect = new Rectangle(
                bounds.X + Scale(10),
                bounds.Y,
                Math.Max(8, (showClose ? close.Left : bounds.Right) - bounds.X - Scale(14)),
                bounds.Height);
            TextRenderer.DrawText(
                g,
                page != null ? page.DisplayTitle : "",
                _tabFont,
                textRect,
                selected ? YiWriteBrowserChromeTheme.Text : YiWriteBrowserChromeTheme.TextMuted,
                TextFormatFlags.EndEllipsis
                | TextFormatFlags.VerticalCenter
                | TextFormatFlags.NoPadding
                | TextFormatFlags.Left);
        }

        private void DrawPlus(Graphics g, Rectangle bounds)
        {
            if (bounds.Width < 4)
            {
                return;
            }

            if (_hotPlus)
            {
                using (var brush = new SolidBrush(YiWriteBrowserChromeTheme.PlusHover))
                {
                    g.FillEllipse(brush, bounds);
                }
            }

            using (var pen = new Pen(YiWriteBrowserChromeTheme.TextMuted, 1.4f * _dpi))
            {
                int cx = bounds.X + bounds.Width / 2;
                int cy = bounds.Y + bounds.Height / 2;
                int arm = Math.Max(4, bounds.Width / 5);
                g.DrawLine(pen, cx - arm, cy, cx + arm, cy);
                g.DrawLine(pen, cx, cy - arm, cx, cy + arm);
            }
        }

        private void LayoutTabs(out List<Rectangle> tabs, out Rectangle plus)
        {
            tabs = new List<Rectangle>();
            int caption = _btnMin.Width + _btnMax.Width + _btnClose.Width;
            int left = Scale(8);
            int plusSize = Scale(28);
            int gap = Scale(4);
            int rightLimit = Width - caption - Scale(8);
            int available = rightLimit - left - plusSize - gap;
            int count = _pages != null ? _pages.Count : 0;
            int tabH = Height - Scale(4);
            int tabY = Scale(4);
            int tabW = count <= 0
                ? Scale(180)
                : Math.Max(Scale(48), Math.Min(Scale(240), available / count));

            int x = left;
            for (int i = 0; i < count; i++)
            {
                if (x + Scale(24) > rightLimit)
                {
                    break;
                }

                int w = Math.Min(tabW, Math.Max(Scale(48), rightLimit - plusSize - gap - x));
                tabs.Add(new Rectangle(x, tabY, w, tabH));
                x += w;
            }

            int plusY = (Height - plusSize) / 2;
            plus = new Rectangle(Math.Min(x + gap, rightLimit - plusSize), plusY, plusSize, plusSize);
        }

        private void HitTest(Point pt, out int tabIndex, out bool close, out bool plus)
        {
            tabIndex = -1;
            close = false;
            plus = false;
            LayoutTabs(out List<Rectangle> tabs, out Rectangle plusRect);
            if (plusRect.Contains(pt))
            {
                plus = true;
                return;
            }

            for (int i = 0; i < tabs.Count; i++)
            {
                if (!tabs[i].Contains(pt))
                {
                    continue;
                }

                tabIndex = i;
                int closeSize = Scale(16);
                int closePad = Scale(8);
                var closeRect = new Rectangle(
                    tabs[i].Right - closePad - closeSize,
                    tabs[i].Y + (tabs[i].Height - closeSize) / 2,
                    closeSize,
                    closeSize);
                close = closeRect.Contains(pt);
                return;
            }
        }

        private static GraphicsPath RoundedTopRect(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (bounds.Width < 2 || bounds.Height < 2)
            {
                path.AddRectangle(bounds);
                return path;
            }

            int r = Math.Max(1, Math.Min(radius, Math.Min(bounds.Width, bounds.Height) / 2));
            path.AddArc(bounds.X, bounds.Y, r * 2, r * 2, 180, 90);
            path.AddArc(bounds.Right - r * 2, bounds.Y, r * 2, r * 2, 270, 90);
            path.AddLine(bounds.Right, bounds.Bottom, bounds.X, bounds.Bottom);
            path.CloseFigure();
            return path;
        }

        private int Scale(int value)
        {
            return (int)Math.Round(value * _dpi);
        }

        [DllImport("user32.dll")]
        private static extern bool ReleaseCapture();

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
    }

    internal sealed class YiWriteBrowserAddressBar : Panel
    {
        private readonly float _dpi;
        private readonly YiWriteBrowserNavButton _back;
        private readonly YiWriteBrowserNavButton _forward;
        private readonly YiWriteBrowserNavButton _reload;
        private readonly TextBox _input;
        private bool _focused;

        public YiWriteBrowserAddressBar(float dpiScale)
        {
            _dpi = dpiScale <= 0 ? 1f : dpiScale;
            DoubleBuffered = true;

            int btn = Scale(32);
            _back = new YiWriteBrowserNavButton(YiWriteBrowserNavKind.Back, _dpi)
            {
                Size = new Size(btn, btn)
            };
            _forward = new YiWriteBrowserNavButton(YiWriteBrowserNavKind.Forward, _dpi)
            {
                Size = new Size(btn, btn)
            };
            _reload = new YiWriteBrowserNavButton(YiWriteBrowserNavKind.Reload, _dpi)
            {
                Size = new Size(btn, btn)
            };
            _back.Click += (s, e) => Raise(BackRequested);
            _forward.Click += (s, e) => Raise(ForwardRequested);
            _reload.Click += (s, e) => Raise(ReloadRequested);

            _input = new TextBox
            {
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular, GraphicsUnit.Point),
                ForeColor = YiWriteBrowserChromeTheme.Text,
                BackColor = YiWriteBrowserChromeTheme.Omnibox,
                TabStop = true
            };
            _input.GotFocus += (s, e) =>
            {
                _focused = true;
                _input.BackColor = YiWriteBrowserChromeTheme.OmniboxFocus;
                Invalidate();
            };
            _input.LostFocus += (s, e) =>
            {
                _focused = false;
                _input.BackColor = YiWriteBrowserChromeTheme.Omnibox;
                Invalidate();
            };

            Controls.Add(_back);
            Controls.Add(_forward);
            Controls.Add(_reload);
            Controls.Add(_input);
            Dock = DockStyle.Top;
            Height = Scale(40);
            BackColor = YiWriteBrowserChromeTheme.Toolbar;
            Padding = new Padding(Scale(8), Scale(4), Scale(12), Scale(4));
            SetNavState(false, false, false);
        }

        public TextBox Input
        {
            get { return _input; }
        }

        public event EventHandler BackRequested;

        public event EventHandler ForwardRequested;

        public event EventHandler ReloadRequested;

        public void SetNavState(bool canGoBack, bool canGoForward, bool canReload)
        {
            if (_back != null)
            {
                _back.Enabled = canGoBack;
            }

            if (_forward != null)
            {
                _forward.Enabled = canGoForward;
            }

            if (_reload != null)
            {
                _reload.Enabled = canReload;
            }
        }

        public void SetSchemeHint(string url)
        {
            Invalidate();
        }

        protected override void OnLayout(LayoutEventArgs levent)
        {
            base.OnLayout(levent);
            if (_input == null)
            {
                return;
            }

            LayoutNavButtons();
            Rectangle field = FieldBounds;
            int icon = Scale(18);
            int textLeft = field.X + Scale(10) + icon + Scale(8);
            int textTop = field.Y + Math.Max(0, (field.Height - _input.PreferredHeight) / 2);
            _input.SetBounds(
                textLeft,
                textTop,
                Math.Max(20, field.Right - Scale(10) - textLeft),
                _input.PreferredHeight);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(YiWriteBrowserChromeTheme.Toolbar))
            {
                g.FillRectangle(bg, ClientRectangle);
            }

            Rectangle field = FieldBounds;
            int radius = Scale(16);
            using (GraphicsPath path = RoundedRect(field, radius))
            using (var fill = new SolidBrush(
                _focused ? YiWriteBrowserChromeTheme.OmniboxFocus : YiWriteBrowserChromeTheme.Omnibox))
            {
                g.FillPath(fill, path);
                if (_focused)
                {
                    using (var pen = new Pen(YiWriteBrowserChromeTheme.OmniboxBorder, 1f))
                    {
                        g.DrawPath(pen, path);
                    }
                }
            }

            DrawLock(g, new Rectangle(field.X + Scale(10), field.Y + (field.Height - Scale(14)) / 2, Scale(14), Scale(14)));
        }

        private void LayoutNavButtons()
        {
            if (_back == null)
            {
                return;
            }

            int btn = Scale(32);
            int gap = Scale(2);
            int x = Padding.Left;
            int y = Math.Max(0, (Height - btn) / 2);
            _back.SetBounds(x, y, btn, btn);
            x += btn + gap;
            _forward.SetBounds(x, y, btn, btn);
            x += btn + gap;
            _reload.SetBounds(x, y, btn, btn);
        }

        private int NavClusterWidth
        {
            get { return Scale(32) * 3 + Scale(2) * 2 + Scale(8); }
        }

        private Rectangle FieldBounds
        {
            get
            {
                int left = Padding.Left + NavClusterWidth;
                return new Rectangle(
                    left,
                    Padding.Top,
                    Math.Max(8, Width - left - Padding.Right),
                    Math.Max(8, Height - Padding.Vertical));
            }
        }

        private static void Raise(EventHandler handler)
        {
            if (handler != null)
            {
                handler(null, EventArgs.Empty);
            }
        }

        private void DrawLock(Graphics g, Rectangle box)
        {
            using (var pen = new Pen(YiWriteBrowserChromeTheme.TextMuted, 1.3f * _dpi))
            {
                var body = new Rectangle(box.X + 2, box.Y + box.Height / 2 - 1, box.Width - 4, box.Height / 2 + 1);
                g.DrawRectangle(pen, body);
                g.DrawArc(
                    pen,
                    box.X + 3,
                    box.Y + 1,
                    box.Width - 6,
                    box.Height - 4,
                    200,
                    140);
            }
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            var path = new GraphicsPath();
            if (bounds.Width < 2 || bounds.Height < 2)
            {
                path.AddRectangle(bounds);
                return path;
            }

            int r = Math.Max(1, Math.Min(radius, Math.Min(bounds.Width, bounds.Height) / 2));
            path.AddArc(bounds.X, bounds.Y, r * 2, r * 2, 180, 90);
            path.AddArc(bounds.Right - r * 2, bounds.Y, r * 2, r * 2, 270, 90);
            path.AddArc(bounds.Right - r * 2, bounds.Bottom - r * 2, r * 2, r * 2, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - r * 2, r * 2, r * 2, 90, 90);
            path.CloseFigure();
            return path;
        }

        private int Scale(int value)
        {
            return (int)Math.Round(value * _dpi);
        }
    }

    internal enum YiWriteBrowserNavKind
    {
        Back,
        Forward,
        Reload
    }

    internal sealed class YiWriteBrowserNavButton : Control
    {
        private readonly float _dpi;
        private readonly YiWriteBrowserNavKind _kind;
        private bool _hover;

        public YiWriteBrowserNavButton(YiWriteBrowserNavKind kind, float dpiScale)
        {
            _kind = kind;
            _dpi = dpiScale <= 0 ? 1f : dpiScale;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);
            TabStop = false;
            BackColor = YiWriteBrowserChromeTheme.Toolbar;
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            Invalidate();
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
            g.Clear(YiWriteBrowserChromeTheme.Toolbar);
            if (Enabled && _hover)
            {
                int pad = Math.Max(2, (int)Math.Round(2 * _dpi));
                var oval = new Rectangle(pad, pad, Width - pad * 2, Height - pad * 2);
                using (var brush = new SolidBrush(YiWriteBrowserChromeTheme.CaptionHover))
                {
                    g.FillEllipse(brush, oval);
                }
            }

            Color icon = Enabled
                ? YiWriteBrowserChromeTheme.TextMuted
                : Color.FromArgb(0xC4, 0xC7, 0xC5);
            float cx = Width / 2f;
            float cy = Height / 2f;
            float s = Math.Min(Width, Height) * 0.16f;
            using (var pen = new Pen(icon, Math.Max(1.4f, 1.5f * _dpi)))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                switch (_kind)
                {
                    case YiWriteBrowserNavKind.Back:
                        g.DrawLines(pen, new[]
                        {
                            new PointF(cx + s * 0.55f, cy - s * 1.15f),
                            new PointF(cx - s * 0.75f, cy),
                            new PointF(cx + s * 0.55f, cy + s * 1.15f)
                        });
                        break;
                    case YiWriteBrowserNavKind.Forward:
                        g.DrawLines(pen, new[]
                        {
                            new PointF(cx - s * 0.55f, cy - s * 1.15f),
                            new PointF(cx + s * 0.75f, cy),
                            new PointF(cx - s * 0.55f, cy + s * 1.15f)
                        });
                        break;
                    case YiWriteBrowserNavKind.Reload:
                        float r = s * 1.28f;
                        var ring = new RectangleF(cx - r, cy - r, r * 2f, r * 2f);
                        const float startDeg = 28f;
                        const float sweepDeg = 302f;
                        g.DrawArc(pen, ring, startDeg, sweepDeg);
                        double endRad = (startDeg + sweepDeg) * Math.PI / 180.0;
                        float ex = cx + r * (float)Math.Cos(endRad);
                        float ey = cy + r * (float)Math.Sin(endRad);
                        g.DrawLine(pen, ex, ey, ex - r * 0.78f, ey);
                        break;
                }
            }
        }
    }

    internal enum YiWriteBrowserCaptionKind
    {
        Minimize,
        Maximize,
        Restore,
        Close
    }

    internal sealed class YiWriteBrowserCaptionButton : Control
    {
        private readonly float _dpi;
        private bool _hover;
        private YiWriteBrowserCaptionKind _kind;

        public YiWriteBrowserCaptionButton(YiWriteBrowserCaptionKind kind, float dpiScale)
        {
            _kind = kind;
            _dpi = dpiScale <= 0 ? 1f : dpiScale;
            SetStyle(
                ControlStyles.AllPaintingInWmPaint
                | ControlStyles.UserPaint
                | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw,
                true);
            TabStop = false;
            BackColor = YiWriteBrowserChromeTheme.Frame;
        }

        public YiWriteBrowserCaptionKind Kind
        {
            get { return _kind; }
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
            Color bg = YiWriteBrowserChromeTheme.Frame;
            if (_hover)
            {
                bg = _kind == YiWriteBrowserCaptionKind.Close
                    ? YiWriteBrowserChromeTheme.CloseHover
                    : YiWriteBrowserChromeTheme.CaptionHover;
            }

            using (var brush = new SolidBrush(bg))
            {
                g.FillRectangle(brush, ClientRectangle);
            }

            Color icon = _hover && _kind == YiWriteBrowserCaptionKind.Close
                ? Color.White
                : YiWriteBrowserChromeTheme.Text;
            using (var pen = new Pen(icon, Math.Max(1f, _dpi)))
            {
                float cx = Width / 2f;
                float cy = Height / 2f;
                float s = Math.Min(Width, Height) * 0.12f;
                switch (_kind)
                {
                    case YiWriteBrowserCaptionKind.Minimize:
                        g.DrawLine(pen, cx - s * 1.2f, cy, cx + s * 1.2f, cy);
                        break;
                    case YiWriteBrowserCaptionKind.Maximize:
                        g.DrawRectangle(pen, cx - s, cy - s, s * 2f, s * 2f);
                        break;
                    case YiWriteBrowserCaptionKind.Restore:
                        float o = s * 0.35f;
                        g.DrawRectangle(pen, cx - s + o, cy - s - o, s * 1.7f, s * 1.7f);
                        using (var erase = new SolidBrush(bg))
                        {
                            g.FillRectangle(erase, cx - s - 0.5f, cy - s + o - 0.5f, s * 1.7f + 1f, s * 1.7f + 1f);
                        }

                        g.DrawRectangle(pen, cx - s, cy - s + o, s * 1.7f, s * 1.7f);
                        break;
                    case YiWriteBrowserCaptionKind.Close:
                        g.DrawLine(pen, cx - s, cy - s, cx + s, cy + s);
                        g.DrawLine(pen, cx + s, cy - s, cx - s, cy + s);
                        break;
                }
            }
        }
    }
}
