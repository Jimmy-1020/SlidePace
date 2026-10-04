using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SlidePace.Setup
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            if (args.Length == 3 && args[0] == "--cleanup")
            {
                Cleanup(args[1], args[2]);
                return;
            }
            bool created;
            using (var mutex = new Mutex(true, "Local\\SlidePaceInstaller-" + Environment.UserName, out created))
            {
                if (!created) { MessageBox.Show("SlidePace 安装程序已经在运行。", "SlidePace"); return; }
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                bool uninstall = args.Contains("--uninstall") || string.Equals(Path.GetFileNameWithoutExtension(Application.ExecutablePath), "Uninstall", StringComparison.OrdinalIgnoreCase);
                Application.Run(new InstallerForm(uninstall));
            }
        }
        private static void Cleanup(string target, string processId)
        {
            string expected = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SlidePace");
            if (!string.Equals(Path.GetFullPath(target).TrimEnd('\\'), Path.GetFullPath(expected).TrimEnd('\\'), StringComparison.OrdinalIgnoreCase)) return;
            int id;
            if (!int.TryParse(processId, out id)) return;
            try { Process.GetProcessById(id).WaitForExit(); } catch (ArgumentException) { }
            string uninstall = Path.Combine(expected, "Uninstall.exe");
            // Only remove the known executable; the main uninstall already removed owned files.
            try
            {
                if (File.Exists(uninstall)) File.Delete(uninstall);
                if (Directory.Exists(expected) && Directory.GetFileSystemEntries(expected).Length == 0) Directory.Delete(expected);
            }
            catch { }
        }
    }

    internal sealed class InstallerForm : Form
    {
        private readonly bool uninstall;
        private readonly Label status = new Label();
        private readonly ProgressBar progress = new ProgressBar();
        private readonly Button action = new Button();
        private readonly CheckBox removeSettings = new CheckBox();
        private readonly InstallOptions options;
        private bool busy;
        private bool successful;
        public InstallerForm(bool removing) : this(removing, new InstallOptions()) { }
        internal InstallerForm(bool removing, InstallOptions installOptions)
        {
            options = installOptions;
            uninstall = removing;
            Text = removing ? "卸载 SlidePace" : "安装 SlidePace";
            ClientSize = new Size(590, 355);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = Color.FromArgb(22, 29, 39);
            ForeColor = Color.FromArgb(229, 237, 247);
            Font = new Font("Microsoft YaHei UI", 10);
            var title = new Label { Text = "SlidePace", Font = new Font("Segoe UI", 28, FontStyle.Bold), Location = new Point(28, 24), Size = new Size(450, 55) };
            var subtitle = new Label { Text = "PowerPoint 放映计时插件 · 1.0", Location = new Point(32, 84), AutoSize = true, ForeColor = Color.FromArgb(147, 193, 234) };
            status.SetBounds(32, 130, 526, 84);
            status.Text = removing ? "移除 PowerPoint 加载项和程序文件。" : "正在准备安装…";
            progress.SetBounds(32, 224, 526, 15);
            removeSettings.Text = "同时删除个人设置";
            removeSettings.SetBounds(32, 261, 240, 32);
            removeSettings.Visible = removing;
            action.SetBounds(424, 271, 134, 40);
            action.Text = removing ? "卸载" : "安装中…";
            action.BackColor = Color.FromArgb(68, 181, 160);
            action.ForeColor = Color.FromArgb(15, 28, 36);
            action.FlatStyle = FlatStyle.Flat;
            action.FlatAppearance.BorderSize = 0;
            action.Click += delegate { if (successful) Close(); else RunOperation(); };
            Controls.AddRange(new Control[] { title, subtitle, status, progress, removeSettings, action });
            FormClosing += delegate(object sender, FormClosingEventArgs args) { if (busy) args.Cancel = true; };
            Shown += delegate { if (!uninstall) RunOperation(); };
            try { Icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath); } catch { }
        }
        private async void RunOperation()
        {
            if (busy) return;
            busy = true;
            action.Enabled = false;
            removeSettings.Enabled = false;
            bool deleteSettings = removeSettings.Checked;
            try
            {
                var engine = new InstallerEngine(options, delegate(string text, int value)
                {
                    BeginInvoke((Action)delegate { status.Text = text; progress.Value = value; });
                });
                await Task.Run(delegate
                {
                    if (uninstall) engine.Uninstall(deleteSettings);
                    else engine.Install();
                });
                successful = true;
                action.Text = "完成";
                if (uninstall) ScheduleCleanup();
            }
            catch (Exception error)
            {
                status.Text = error.Message;
                progress.Value = 0;
                action.Text = "重试";
            }
            finally { busy = false; action.Enabled = true; removeSettings.Enabled = !successful; }
        }
        private void ScheduleCleanup()
        {
            // A separate copy waits for this process to close, so Windows can delete the in-use uninstaller.
            if (!string.Equals(Path.GetDirectoryName(Application.ExecutablePath), Path.GetFullPath(options.Root), StringComparison.OrdinalIgnoreCase)) return;
            string helper = Path.Combine(Path.GetTempPath(), "SlidePace-Cleanup-" + Guid.NewGuid().ToString("N") + ".exe");
            try
            {
                File.Copy(Application.ExecutablePath, helper);
                Process.Start(new ProcessStartInfo(helper, "--cleanup \"" + options.Root + "\" " + Process.GetCurrentProcess().Id)
                { UseShellExecute = false, CreateNoWindow = true, WindowStyle = ProcessWindowStyle.Hidden });
            }
            catch { status.Text = "插件已卸载；程序文件将在关闭后可手动移除。"; }
        }
    }
}
