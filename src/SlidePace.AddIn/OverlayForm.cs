using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SlidePace
{
    internal sealed class NativeOwner : IWin32Window
    {
        public IntPtr Handle { get; private set; }
        public NativeOwner(IntPtr handle) { Handle = handle; }
    }

    internal sealed class OverlayForm : Form
    {
        private readonly Label collapsedLabel = new Label();
        private sealed class DigitsLabel : Label
        {
            internal Rectangle InkBounds;
            protected override void OnPaint(PaintEventArgs args)
            {
                Point origin = new Point((Width - InkBounds.Width) / 2 - InkBounds.Left,
                    (Height - InkBounds.Height) / 2 - InkBounds.Top);
                TextRenderer.DrawText(args.Graphics, Text, Font, origin, ForeColor,
                    TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
            }
        }
        private readonly DigitsLabel digits = new DigitsLabel();
        private readonly Button start = new Button();
        private readonly Button resume = new Button();
        private readonly Button reset = new Button();
        private readonly Panel controls = new Panel();
        private TimerSnapshot snapshot;
        private UserSettings settings;
        private bool dragging;
        private Point dragAnchor;
        private Point originalLocation;
        private int profile = -1;
        private Font numberFont;
        private string fontName;
        private int fontProfile = -1;
        private bool controlsExpanded;
        private bool controlsAbove;
        private bool layingOut;
        private int expansionHeight;
        private int compactHeight;
        private string layoutText;
        public Rectangle DisplayArea { get; set; }
        public Rectangle AreaBounds { get { return DisplayArea.IsEmpty ? Screen.FromControl(this).Bounds : DisplayArea; } }
        public string DeviceName { get; private set; }
        public bool IsPresenter { get; private set; }
        public bool Collapsed { get; private set; }
        public bool ControlsVisible { get { return controls.Visible; } }
        public string DisplayText { get { return digits.Text; } }
        public Color DisplayColor { get { return digits.ForeColor; } }
        public string DisplayFontName { get { return digits.Font.FontFamily.Name; } }
        public Rectangle CompactBounds
        {
            get { return new Rectangle(Left, Top + (controlsExpanded && controlsAbove ? expansionHeight : 0), Width, compactHeight); }
        }
        public event Action StartPauseRequested;
        public event Action ContinueRequested;
        public event Action ResetRequested;
        public event Action SettingsRequested;
        public event Action CollapseRequested;
        public event Action<TimerMode> ModeRequested;
        public event Action<OverlayForm> DragFinished;

        public OverlayForm(string device, bool presenter)
        {
            DeviceName = device;
            IsPresenter = presenter;
            Text = "SlidePace 计时";
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = false;
            // Leave WinForms TopMost false: its focus paths activate the form.
            // KeepAboveShow promotes the native window without activation.
            AutoScaleMode = AutoScaleMode.None;
            DoubleBuffered = true;
            BackColor = Color.FromArgb(22, 29, 39);
            Opacity = 0.82;
            Font = new Font("Microsoft YaHei UI", 9);
            collapsedLabel.Text = "点击展开计时";
            collapsedLabel.ForeColor = Color.FromArgb(209, 219, 231);
            collapsedLabel.TextAlign = ContentAlignment.MiddleCenter;
            collapsedLabel.Visible = false;
            digits.ForeColor = Color.White;
            digits.TextAlign = ContentAlignment.MiddleCenter;
            digits.AccessibleName = "时间";
            controls.BackColor = BackColor;
            controls.Visible = false;
            ConfigureButton(start, "开始");
            ConfigureButton(resume, "继续");
            ConfigureButton(reset, "重置");
            start.Click += delegate { if (StartPauseRequested != null) StartPauseRequested(); };
            resume.Click += delegate { if (ContinueRequested != null) ContinueRequested(); };
            reset.Click += delegate { if (ResetRequested != null) ResetRequested(); };
            controls.Controls.AddRange(new Control[] { start, resume, reset });
            Controls.AddRange(new Control[] { collapsedLabel, digits, controls });
            AttachDrag(this);
            AttachDrag(collapsedLabel);
            AttachDrag(digits);
            foreach (Control target in new Control[] { this, collapsedLabel, digits, controls, start, resume, reset })
            {
                target.MouseEnter += delegate { RefreshHover(); };
                target.MouseLeave += delegate { RefreshHover(); };
            }
            var menu = new ContextMenuStrip();
            menu.Items.Add("计时器设置…", null, delegate { if (SettingsRequested != null) SettingsRequested(); });
            menu.Items.Add("收起／展开框体", null, delegate { if (CollapseRequested != null) CollapseRequested(); });
            menu.Items.Add(new ToolStripSeparator());
            foreach (TimerMode mode in Enum.GetValues(typeof(TimerMode)))
            {
                if (mode == TimerMode.None) continue;
                TimerMode choice = mode;
                var item = new ToolStripMenuItem(TimerEngine.ModeName(mode));
                item.Tag = mode;
                item.Click += delegate { if (ModeRequested != null) ModeRequested(choice); };
                menu.Items.Add(item);
            }
            ContextMenuStrip = menu;
        }

        protected override bool ShowWithoutActivation { get { return true; } }
        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);

        internal void KeepAboveShow()
        {
            // PowerPoint adjusts its full-screen Z order after SlideShowBegin.
            // Reassert the native order after display without moving or activating.
            if (Visible) SetWindowPos(Handle, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002 | 0x0010 | 0x4000);
        }
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams parameters = base.CreateParams;
                parameters.ExStyle |= 0x08000000 | 0x00000080; // No activation, tool window.
                return parameters;
            }
        }
        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0x0021) { message.Result = new IntPtr(3); return; } // MA_NOACTIVATE.
            base.WndProc(ref message);
        }

        private void ConfigureButton(Button button, string label)
        {
            button.Text = label;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = Color.FromArgb(42, 54, 70);
            button.ForeColor = Color.FromArgb(230, 238, 248);
            button.TabStop = false;
            button.Cursor = Cursors.Hand;
        }

        private int ScalePixel(int value) { return (int)Math.Round(value * DeviceDpi / 96.0); }

        public void ApplySettings(UserSettings value, bool collapsed)
        {
            settings = value;
            bool changed = profile != value.FontSize || fontName != value.NumberFontName || Collapsed != collapsed;
            profile = value.FontSize;
            Collapsed = collapsed;
            if (changed || numberFont == null) LayoutCard(false);
        }

        private void LayoutCard(bool showControls)
        {
            if (layingOut) return;
            layingOut = true;
            try
            {
                Rectangle previousCompact = CompactBounds;
                Point compactLocation = previousCompact.Location;
                bool wasExpanded = controlsExpanded;
                if (numberFont == null || fontProfile != profile || fontName != settings.NumberFontName)
                {
                    Font previous = numberFont;
                    using (var family = new FontFamily(settings.NumberFontName))
                        numberFont = new Font(family, profile == 0 ? 29 : profile == 2 ? 45 : 36,
                            family.IsStyleAvailable(FontStyle.Bold) ? FontStyle.Bold : FontStyle.Regular);
                    digits.Font = numberFont;
                    if (previous != null) previous.Dispose();
                    fontProfile = profile;
                    fontName = settings.NumberFontName;
                }
                int width = ScalePixel(150);
                compactHeight = ScalePixel(Collapsed ? 38 : profile == 0 ? 68 : profile == 2 ? 98 : 80);
                if (!Collapsed)
                {
                    layoutText = string.IsNullOrEmpty(digits.Text) ? "00:00:00" : digits.Text;
                    Rectangle ink = MeasureDigitInk(layoutText);
                    digits.InkBounds = ink;
                    compactHeight = Math.Max(compactHeight, ink.Height + ScalePixel(16));
                    width = ink.Width + compactHeight - ink.Height;
                }
                Rectangle screen = DisplayArea.IsEmpty ? Screen.FromPoint(new Point(compactLocation.X + width / 2,
                    compactLocation.Y + compactHeight / 2)).Bounds : DisplayArea;
                width = Math.Min(width, screen.Width);
                OverlayPosition position = IsPresenter ? settings.PresenterPosition : settings.AudiencePosition;
                if (previousCompact.Height > 0 && (position == OverlayPosition.TopRight || position == OverlayPosition.BottomRight))
                    compactLocation.X = previousCompact.Right - width;
                controlsExpanded = showControls && !Collapsed;
                expansionHeight = controlsExpanded ? ScalePixel(40) : 0;
                if (!wasExpanded || !controlsExpanded)
                    controlsAbove = controlsExpanded && compactLocation.Y + compactHeight + expansionHeight > screen.Bottom;
                ClientSize = new Size(width, compactHeight + expansionHeight);
                int digitOffset = controlsExpanded && controlsAbove ? expansionHeight : 0;
                int padding = ScalePixel(8);
                digits.SetBounds(padding, digitOffset + padding, width - padding * 2, compactHeight - padding * 2);
                collapsedLabel.SetBounds(padding, padding, width - padding * 2, compactHeight - padding * 2);
                controls.SetBounds(ScalePixel(16), controlsAbove ? padding : compactHeight + ScalePixel(1),
                    width - ScalePixel(32), ScalePixel(31));
                int gap = ScalePixel(6);
                int buttonWidth = (controls.Width - gap * 2) / 3;
                start.SetBounds(0, 0, buttonWidth, controls.Height);
                resume.SetBounds(buttonWidth + gap, 0, buttonWidth, controls.Height);
                reset.SetBounds((buttonWidth + gap) * 2, 0, buttonWidth, controls.Height);
                digits.Visible = !Collapsed;
                collapsedLabel.Visible = Collapsed;
                controls.Visible = controlsExpanded;
                Location = Clamp(new Point(compactLocation.X, compactLocation.Y - digitOffset), screen, Size);
                using (GraphicsPath path = RoundedRectangle(new Rectangle(0, 0, Width, Height), ScalePixel(14)))
                {
                    Region oldRegion = Region;
                    Region = new Region(path);
                    if (oldRegion != null) oldRegion.Dispose();
                }
                Invalidate();
            }
            finally { layingOut = false; }
        }

        protected override void OnDpiChanged(DpiChangedEventArgs args)
        {
            base.OnDpiChanged(args);
            LayoutCard(controlsExpanded);
            Location = Clamp(Location, AreaBounds, Size);
        }

        private Rectangle MeasureDigitInk(string text)
        {
            using (Graphics display = CreateGraphics())
            {
                Size measure = TextRenderer.MeasureText(display, text, numberFont, Size.Empty,
                    TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
                using (var image = new Bitmap(measure.Width + 16, measure.Height + 16, PixelFormat.Format32bppArgb))
                {
                    image.SetResolution(display.DpiX, display.DpiY);
                    using (Graphics graphics = Graphics.FromImage(image))
                    {
                        graphics.Clear(Color.Black);
                        TextRenderer.DrawText(graphics, text, numberFont, new Point(8, 8), Color.White,
                            TextFormatFlags.NoPadding | TextFormatFlags.SingleLine);
                    }
                    // Font line metrics include unused ascent/descent space. Measure
                    // the rendered strokes so the visible margins match on all sides.
                    int left = image.Width, top = image.Height, right = -1, bottom = -1;
                    BitmapData data = image.LockBits(new Rectangle(Point.Empty, image.Size), ImageLockMode.ReadOnly, image.PixelFormat);
                    try
                    {
                        byte[] pixels = new byte[data.Stride * image.Height];
                        Marshal.Copy(data.Scan0, pixels, 0, pixels.Length);
                        for (int y = 0; y < image.Height; y++)
                            for (int x = 0; x < image.Width; x++)
                            {
                                int offset = y * data.Stride + x * 4;
                                if (Math.Max(pixels[offset], Math.Max(pixels[offset + 1], pixels[offset + 2])) <= 24) continue;
                                left = Math.Min(left, x); right = Math.Max(right, x);
                                top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                            }
                    }
                    finally { image.UnlockBits(data); }
                    return right < left ? new Rectangle(Point.Empty, measure) :
                        new Rectangle(left - 8, top - 8, right - left + 1, bottom - top + 1);
                }
            }
        }

        public void Render(TimerSnapshot value)
        {
            snapshot = value;
            digits.Text = value.Text;
            if (layoutText != value.Text && !Collapsed) LayoutCard(controlsExpanded);
            digits.AccessibleDescription = value.Status;
            Color color = value.IsOvertime ? settings.GetOvertimeColor() : Color.White;
            digits.ForeColor = color;
            start.Text = value.IsRunning ? "暂停" : "开始";
            start.Enabled = value.IsRunning || !value.HasStarted;
            resume.Enabled = value.HasStarted && !value.IsRunning && !value.IsCompleted;
            reset.Enabled = true;
            foreach (ToolStripItem entry in ContextMenuStrip.Items)
            {
                var item = entry as ToolStripMenuItem;
                if (item != null && item.Tag is TimerMode) item.Checked = (TimerMode)item.Tag == value.Mode;
            }
            RefreshHover();
        }

        public void RefreshHover()
        {
            if (layingOut || settings == null) return;
            bool hovered = dragging || Bounds.Contains(Cursor.Position) || (ContextMenuStrip != null && ContextMenuStrip.Visible);
            bool showControls = !Collapsed && snapshot != null && snapshot.Mode != TimerMode.Clock && hovered;
            if (showControls != controlsExpanded) LayoutCard(showControls);
        }

        private void AttachDrag(Control target)
        {
            target.MouseDown += delegate(object sender, MouseEventArgs args)
            {
                if (args.Button != MouseButtons.Left) return;
                dragging = true;
                dragAnchor = Cursor.Position;
                originalLocation = Location;
                target.Capture = true;
            };
            target.MouseMove += delegate
            {
                if (!dragging) return;
                Point pointer = Cursor.Position;
                Rectangle area = AreaBounds;
                Location = Clamp(new Point(originalLocation.X + pointer.X - dragAnchor.X,
                    originalLocation.Y + pointer.Y - dragAnchor.Y), area, Size);
            };
            target.MouseUp += delegate
            {
                if (!dragging) return;
                dragging = false;
                target.Capture = false;
                if (Collapsed && Math.Abs(Cursor.Position.X - dragAnchor.X) + Math.Abs(Cursor.Position.Y - dragAnchor.Y) < 4)
                {
                    if (CollapseRequested != null) CollapseRequested();
                }
                else if (DragFinished != null) DragFinished(this);
            };
        }

        public static Point Clamp(Point location, Rectangle screen, Size size)
        {
            return new Point(Math.Max(screen.Left, Math.Min(location.X, screen.Right - size.Width)),
                Math.Max(screen.Top, Math.Min(location.Y, screen.Bottom - size.Height)));
        }

        private static GraphicsPath RoundedRectangle(Rectangle rectangle, int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;
            path.AddArc(rectangle.Left, rectangle.Top, diameter, diameter, 180, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Top, diameter, diameter, 270, 90);
            path.AddArc(rectangle.Right - diameter, rectangle.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rectangle.Left, rectangle.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }
        protected override void OnPaint(PaintEventArgs args)
        {
            base.OnPaint(args);
            args.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = RoundedRectangle(new Rectangle(1, 1, Width - 2, Height - 2), ScalePixel(14)))
            using (var pen = new Pen(Color.FromArgb(67, 83, 102))) args.Graphics.DrawPath(pen, path);
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (numberFont != null) { numberFont.Dispose(); numberFont = null; }
                if (ContextMenuStrip != null) ContextMenuStrip.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
