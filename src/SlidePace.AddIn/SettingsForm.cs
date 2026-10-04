using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SlidePace
{
    internal sealed class SettingsForm : Form
    {
        private readonly UserSettings original;
        private readonly ComboBox mode = new ComboBox();
        private readonly NumericUpDown hours = new NumericUpDown();
        private readonly NumericUpDown minutes = new NumericUpDown();
        private readonly NumericUpDown seconds = new NumericUpDown();
        private readonly ComboBox size = new ComboBox();
        private readonly ComboBox numberFont = new ComboBox();
        private readonly Label fontPreview = new Label();
        private Font previewFont;
        private readonly ComboBox presenter = new ComboBox();
        private readonly ComboBox audience = new ComboBox();
        private readonly ComboBox presenterPosition = new ComboBox();
        private readonly ComboBox audiencePosition = new ComboBox();
        private readonly Button colorButton = new Button();
        private readonly List<string> devices = new List<string>();
        private Color overtimeColor;
        public UserSettings Result { get; private set; }

        public SettingsForm(UserSettings settings, bool countdownRunning)
        {
            original = settings;
            Text = "SlidePace · 框体设置";
            Font = new Font("Microsoft YaHei UI", 9);
            BackColor = Color.FromArgb(246, 248, 251);
            ForeColor = Color.FromArgb(27, 40, 56);
            ClientSize = new Size(620, 620);
            MinimumSize = new Size(636, 659);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Dpi;
            var table = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(24), ColumnCount = 2, RowCount = 13 };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 145));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int index = 0; index < 12; index++) table.RowStyles.Add(new RowStyle(SizeType.Absolute, index == 0 ? 45 : index == 7 ? 52 : 38));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            var title = new Label { Text = "放映自动开始 · 双方同步显示", AutoSize = true, Font = new Font(Font.FontFamily, 12, FontStyle.Bold) };
            table.Controls.Add(title, 0, 0);
            table.SetColumnSpan(title, 2);
            mode.Items.AddRange(new object[] { "不显示（全部不选）", "顺计时", "倒计时", "系统时间" });
            mode.SelectedIndex = settings.Mode;
            AddRow(table, 1, "计时模式", mode);
            var duration = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false };
            SetNumber(hours, 23, settings.CountdownSeconds / 3600);
            SetNumber(minutes, 59, settings.CountdownSeconds / 60 % 60);
            SetNumber(seconds, 59, settings.CountdownSeconds % 60);
            duration.Controls.AddRange(new Control[] { hours, new Label { Text = "时", AutoSize = true, Margin = new Padding(1, 6, 6, 0) },
                minutes, new Label { Text = "分", AutoSize = true, Margin = new Padding(1, 6, 6, 0) }, seconds,
                new Label { Text = "秒", AutoSize = true, Margin = new Padding(1, 6, 0, 0) } });
            duration.Enabled = !countdownRunning;
            AddRow(table, 2, "倒计时时长", duration);
            var hint = new Label { Text = countdownRunning ? "倒计时运行中：请先暂停，再修改时长。" : "归零后继续显示超时时间。", AutoSize = true, ForeColor = Color.FromArgb(92, 111, 134) };
            table.Controls.Add(hint, 1, 3);
            overtimeColor = settings.GetOvertimeColor();
            colorButton.Text = "选择超时颜色…";
            colorButton.BackColor = overtimeColor;
            colorButton.ForeColor = overtimeColor.GetBrightness() < 0.5 ? Color.White : Color.Black;
            colorButton.Click += delegate
            {
                using (var dialog = new ColorDialog { Color = overtimeColor, FullOpen = true })
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        overtimeColor = dialog.Color;
                        colorButton.BackColor = overtimeColor;
                        colorButton.ForeColor = overtimeColor.GetBrightness() < 0.5 ? Color.White : Color.Black;
                    }
            };
            AddRow(table, 4, "归零后的数字颜色", colorButton);
            size.Items.AddRange(new object[] { "小", "中", "大" });
            size.SelectedIndex = settings.FontSize;
            AddRow(table, 5, "数字大小", size);
            numberFont.Name = "NumberFont";
            numberFont.AccessibleName = "数字字体";
            foreach (FontFamily family in FontFamily.Families.OrderBy(delegate(FontFamily value) { return value.Name; }))
            {
                using (family)
                    if (family.IsStyleAvailable(FontStyle.Regular) || family.IsStyleAvailable(FontStyle.Bold))
                        numberFont.Items.Add(family.Name);
            }
            numberFont.SelectedItem = settings.NumberFontName;
            if (numberFont.SelectedIndex < 0) numberFont.SelectedItem = "Consolas";
            fontPreview.Text = "00:00:00   -00:00:01";
            fontPreview.TextAlign = ContentAlignment.MiddleLeft;
            numberFont.SelectedIndexChanged += delegate { UpdateFontPreview(); };
            AddRow(table, 6, "数字字体", numberFont);
            AddRow(table, 7, "字体预览", fontPreview);
            UpdateFontPreview();
            devices.Add("");
            presenter.Items.Add("自动识别");
            audience.Items.Add("自动识别");
            int screenIndex = 1;
            foreach (Screen screen in Screen.AllScreens)
            {
                devices.Add(screen.DeviceName);
                string label = "屏幕 " + screenIndex++ + (screen.Primary ? "（主屏）" : "") + " · " + screen.Bounds.Width + " × " + screen.Bounds.Height;
                presenter.Items.Add(label);
                audience.Items.Add(label);
            }
            presenter.SelectedIndex = Math.Max(0, devices.IndexOf(settings.PresenterDevice));
            audience.SelectedIndex = Math.Max(0, devices.IndexOf(settings.AudienceDevice));
            AddRow(table, 8, "演讲者屏幕", presenter);
            AddRow(table, 9, "观众放映屏幕", audience);
            FillPositions(presenterPosition, settings.PresenterPosition);
            FillPositions(audiencePosition, settings.AudiencePosition);
            AddRow(table, 10, "演讲者框体位置", presenterPosition);
            AddRow(table, 11, "观众框体位置", audiencePosition);
            var footer = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(0, 10, 0, 0) };
            var save = new Button { Text = "保存设置", Width = 112, Height = 34, BackColor = Color.FromArgb(35, 93, 179), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            var cancel = new Button { Text = "取消", Width = 90, Height = 34, DialogResult = DialogResult.Cancel };
            save.Click += Save;
            footer.Controls.AddRange(new Control[] { save, cancel });
            table.Controls.Add(footer, 0, 12);
            table.SetColumnSpan(footer, 2);
            Controls.Add(table);
            AcceptButton = save;
            CancelButton = cancel;
        }

        private static void AddRow(TableLayoutPanel table, int row, string text, Control control)
        {
            var label = new Label { Text = text, AutoSize = true, Anchor = AnchorStyles.Left };
            control.Dock = DockStyle.Fill;
            var combo = control as ComboBox;
            if (combo != null) combo.DropDownStyle = ComboBoxStyle.DropDownList;
            table.Controls.Add(label, 0, row);
            table.Controls.Add(control, 1, row);
        }
        private static void SetNumber(NumericUpDown control, int max, int value)
        {
            control.Minimum = 0;
            control.Maximum = max;
            control.Value = value;
            control.Width = 65;
        }
        private static void FillPositions(ComboBox control, OverlayPosition position)
        {
            control.Items.AddRange(new object[] { "右上方（默认）", "左上方", "右下方", "左下方", "自定义（拖动框体设置）" });
            control.SelectedIndex = (int)position;
        }
        private void UpdateFontPreview()
        {
            Font previous = previewFont;
            using (var family = new FontFamily((string)numberFont.SelectedItem ?? "Consolas"))
                previewFont = new Font(family, 16, family.IsStyleAvailable(FontStyle.Bold) ? FontStyle.Bold : FontStyle.Regular);
            fontPreview.Font = previewFont;
            if (previous != null) previous.Dispose();
        }
        protected override void Dispose(bool disposing)
        {
            if (disposing && previewFont != null) { previewFont.Dispose(); previewFont = null; }
            base.Dispose(disposing);
        }
        private void Save(object sender, EventArgs args)
        {
            int duration = (int)(hours.Value * 3600 + minutes.Value * 60 + seconds.Value);
            if (duration < 1)
            {
                MessageBox.Show(this, "倒计时时长至少为 1 秒。", "SlidePace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Result = new UserSettings
            {
                Mode = mode.SelectedIndex,
                CountdownSeconds = duration,
                FontSize = size.SelectedIndex,
                NumberFontName = (string)numberFont.SelectedItem ?? "Consolas",
                OvertimeColor = ColorTranslator.ToHtml(overtimeColor),
                PresenterDevice = devices[presenter.SelectedIndex],
                AudienceDevice = devices[audience.SelectedIndex],
                PresenterPosition = (OverlayPosition)presenterPosition.SelectedIndex,
                AudiencePosition = (OverlayPosition)audiencePosition.SelectedIndex,
                ShowOverlay = original.ShowOverlay,
                Locations = original.Locations.Select(delegate(SavedLocation item) { return new SavedLocation { Device = item.Device, X = item.X, Y = item.Y }; }).ToList()
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
