using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using SlidePace.Setup;

internal static partial class Program
{
    private static void SetupInterface()
    {
        OfficeEnvironment environment = OfficeEnvironment.Detect();
        Check(File.Exists(environment.Executable), "actual installer preflight detects x64 PowerPoint and .NET Framework 4.8");
        Console.WriteLine("Installer preflight Office file version: " + environment.Version);
        string prefix = @"Software\SlidePace.Tests\" + Guid.NewGuid().ToString("N");
        var options = new InstallOptions
        {
            Root = Path.Combine(root, "setup-ui-" + Guid.NewGuid().ToString("N")),
            SettingsRoot = Path.Combine(root, "setup-ui-settings"),
            RegistryPrefix = prefix,
            TestMode = true
        };
        try
        {
            using (var form = new InstallerForm(false, options))
            {
                form.Show();
                Until(delegate { return form.Controls.OfType<Button>().Single().Text == "完成"; }, 10000, "installer UI auto installs after show");
                Equal(form.Controls.OfType<ProgressBar>().Single().Value, 100, "installer progress reaches complete");
                Check(File.Exists(Path.Combine(options.Root, "Uninstall.exe")), "UI deployed uninstaller");
                Pump(150);
                using (var image = new Bitmap(form.Width, form.Height))
                using (Graphics graphics = Graphics.FromImage(image))
                {
                    graphics.CopyFromScreen(form.Location, Point.Empty, form.Size);
                    image.Save(Path.Combine(root, "installer.png"), ImageFormat.Png);
                }
                form.Close();
            }
            using (var form = new InstallerForm(true, options))
            {
                form.Show();
                form.Controls.OfType<Button>().Single().PerformClick();
                Until(delegate { return form.Controls.OfType<Button>().Single().Text == "完成"; }, 10000, "uninstaller UI completes");
                Equal(Directory.GetFileSystemEntries(options.Root).Length, 0, "uninstaller UI removes owned deployment");
                form.Close();
            }
        }
        finally
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)) registry.DeleteSubKeyTree(prefix, false);
        }
    }
}
