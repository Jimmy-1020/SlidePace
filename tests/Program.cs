using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Win32;
using SlidePace;
using SlidePace.Setup;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal sealed class FakeTime : ITimeSource
{
    public double Seconds;
    public DateTime Now = new DateTime(2026, 10, 4, 23, 59, 59);
    public double MonotonicSeconds { get { return Seconds; } }
    public DateTime LocalTime { get { return Now; } }
}

internal static partial class Program
{
    private static int assertions;
    private static string root;
    [STAThread]
    private static int Main(string[] args)
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        root = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.GetFullPath("artifacts/tests");
        Directory.CreateDirectory(root);
        try
        {
            string suite = args.Length == 0 ? "--core" : args[0];
            if (suite == "--core") Core();
            else if (suite == "--installer") Installer();
            else if (suite == "--ui") UserInterface();
            else if (suite == "--office") OfficeIntegration();
            else if (suite == "--host-load") HostLoading();
            else if (suite == "--com-abi") NativeComContract();
            else if (suite == "--setup-ui") SetupInterface();
            else throw new ArgumentException("Unknown suite: " + suite);
            Console.WriteLine("PASS " + suite + " (" + assertions + " assertions)");
            return 0;
        }
        catch (Exception error) { Console.Error.WriteLine(error); return 1; }
    }
    private static void Check(bool condition, string message)
    {
        assertions++;
        if (!condition) throw new InvalidOperationException("FAIL: " + message);
    }
    private static void Equal<T>(T actual, T expected, string message) { Check(Equals(actual, expected), message + ": actual=" + actual + ", expected=" + expected); }
    private static void Pump(int milliseconds)
    {
        var watch = Stopwatch.StartNew();
        while (watch.ElapsedMilliseconds < milliseconds) { Application.DoEvents(); Thread.Sleep(10); }
        Application.DoEvents();
    }
    private static void Until(Func<bool> predicate, int timeout, string message)
    {
        var watch = Stopwatch.StartNew();
        while (!predicate() && watch.ElapsedMilliseconds < timeout) Pump(20);
        Check(predicate(), message);
    }
    private static void Core()
    {
        var time = new FakeTime();
        var timer = new TimerEngine(time);
        timer.BeginSlideShow();
        Equal(timer.Mode, TimerMode.None, "initial selection is none");
        Check(!timer.Snapshot().IsRunning, "none never starts");
        timer.SelectMode(TimerMode.CountUp);
        Check(timer.Snapshot().IsRunning, "selection during show auto starts");
        time.Seconds = 61.25;
        Equal(timer.Snapshot().Text, "00:01:01", "count up uses elapsed time");
        timer.Pause();
        time.Seconds += 500;
        Equal(timer.Snapshot().Text, "00:01:01", "pause excludes time");
        timer.Start();
        timer.Start();
        time.Seconds += 3599;
        Equal(timer.Snapshot().Text, "01:01:00", "repeat start is idempotent and hours roll over");
        time.Now = time.Now.AddDays(-5);
        Equal(timer.Snapshot().Text, "01:01:00", "system clock changes do not change elapsed");
        timer.EndSlideShow();
        time.Seconds += 100;
        Equal(timer.Snapshot().Text, "01:01:00", "show exit pauses");
        timer.BeginSlideShow();
        time.Seconds += 1;
        Equal(timer.Snapshot().Text, "01:01:01", "second show auto resumes");
        timer.Reset();
        time.Seconds += 1;
        Equal(timer.Snapshot().Text, "00:00:00", "reset does not auto start on refresh");
        Check(!timer.Snapshot().HasStarted, "reset clears started state");
        timer.SetCountdownSeconds(5);
        timer.SelectMode(TimerMode.CountDown);
        double start = time.Seconds;
        time.Seconds = start + 4.999;
        Equal(timer.Snapshot().Text, "00:00:01", "positive fraction rounds up");
        Check(!timer.Snapshot().IsOvertime, "not early overtime");
        time.Seconds = start + 5;
        Equal(timer.Snapshot().Text, "00:00:00", "exact boundary is zero");
        Check(timer.Snapshot().IsRunning && timer.Snapshot().IsOvertime, "zero continues running in overtime");
        time.Seconds += 0.4;
        Equal(timer.Snapshot().Text, "00:00:00", "no negative zero");
        time.Seconds = start + 8;
        Equal(timer.Snapshot().Text, "-00:00:03", "overtime grows");
        timer.Pause();
        time.Seconds += 20;
        Equal(timer.Snapshot().Text, "-00:00:03", "overtime pause");
        timer.Start();
        time.Seconds += 2;
        Equal(timer.Snapshot().Text, "-00:00:05", "overtime resume");
        timer.SelectMode(TimerMode.None);
        Check(!timer.Snapshot().IsRunning, "none pauses");
        time.Seconds += 50;
        timer.SelectMode(TimerMode.CountDown);
        Equal(timer.Snapshot().Text, "-00:00:05", "cancel preserves overtime");
        Check(timer.Snapshot().IsRunning, "return during show resumes");
        timer.Reset();
        Equal(timer.Snapshot().Text, "00:00:05", "countdown reset");
        Check(!timer.Snapshot().IsOvertime && !timer.Snapshot().IsRunning, "reset clears overtime and running");
        timer.SetCountdownSeconds(1);
        timer.Start();
        bool rejected = false;
        try { timer.SetCountdownSeconds(10); } catch (InvalidOperationException) { rejected = true; }
        Check(rejected, "running duration change rejected");
        timer.SelectMode(TimerMode.Clock);
        time.Now = new DateTime(2026, 10, 4, 23, 59, 59);
        Equal(timer.Snapshot().Text, "23:59:59", "clock 24h");
        time.Now = time.Now.AddSeconds(1);
        Equal(timer.Snapshot().Text, "00:00:00", "clock midnight");
        Check(!timer.Snapshot().IsRunning, "clock has no counter controls");
        int parsed;
        Check(TimerEngine.TryParseDuration("23:59:59", out parsed) && parsed == 86399, "max duration");
        foreach (string invalid in new[] { "", "-00:10:00", "00:00:00", "24:00:00", "00:60:00", "1e2:00:00", "abc", "00:00:-1" })
            Check(!TimerEngine.TryParseDuration(invalid, out parsed), "reject invalid duration " + invalid);
        time = new FakeTime();
        timer = new TimerEngine(time);
        timer.SelectMode(TimerMode.CountUp);
        timer.BeginSlideShow();
        time.Seconds = 1800;
        Equal(timer.Snapshot().Text, "00:30:00", "simulated 30-minute elapsed accuracy");
        time.Seconds = 7200;
        Equal(timer.Snapshot().Text, "02:00:00", "simulated two-hour elapsed accuracy");
        var store = new SettingsStore(Path.Combine(root, "settings"));
        var settings = new UserSettings { Mode = 2, OvertimeColor = "#22CC88", PresenterPosition = OverlayPosition.BottomLeft, NumberFontName = "Arial" };
        store.Save(settings);
        var loaded = store.Load();
        Equal(loaded.Mode, 0, "new session ignores last mode selection");
        Equal(loaded.NumberFontName, "Arial", "selected number font persists");
        Equal(loaded.GetOvertimeColor().ToArgb(), Color.FromArgb(0x22, 0xCC, 0x88).ToArgb(), "custom color persists");
        Equal(loaded.PresenterPosition, OverlayPosition.BottomLeft, "position persists");
        settings.Mode = 0;
        store.Save(settings);
        Equal(store.Load().Mode, 0, "none persists");
        File.WriteAllText(store.FilePath, "not json");
        Equal(store.Load().CountdownSeconds, 600, "corrupt config falls back");
        File.WriteAllText(store.FilePath, "{}");
        Check(store.Load().ShowOverlay && store.Load().FontSize == 1 && store.Load().NumberFontName == "Consolas", "missing configuration fields preserve defaults");
        settings = new UserSettings { Mode = 999, CountdownSeconds = -2, OvertimeColor = "garbage", FontSize = 10, NumberFontName = "SlidePace Missing Font" };
        settings.Validate();
        Equal(settings.Mode, 0, "invalid mode fallback");
        Equal(settings.CountdownSeconds, 600, "invalid duration fallback");
        Equal(settings.GetOvertimeColor().ToArgb(), Color.Red.ToArgb(), "invalid color fallback");
        Equal(settings.NumberFontName, "Consolas", "unavailable number font falls back");
        var plugin = new PowerPointAddIn(Path.Combine(root, "ribbon"));
        XDocument ribbon = XDocument.Parse(plugin.GetCustomUI("Microsoft.PowerPoint.Presentation"));
        Check(!ribbon.Descendants().Any(delegate(XElement element) { return (string)element.Attribute("id") == "ModeNone" || (string)element.Attribute("id") == "SlidePaceControls"; }), "ribbon has no clear-mode button or timer control group");
        foreach (XAttribute attribute in ribbon.Descendants().Attributes().Where(delegate(XAttribute item) { return item.Name.LocalName.StartsWith("on") || item.Name.LocalName.StartsWith("get"); }))
            Check(typeof(PowerPointAddIn).GetMethod(attribute.Value) != null, "ribbon callback exists: " + attribute.Value);
        Check(typeof(PowerPointAddIn).GetInterfaces().Any(delegate(Type type) { return type.GUID == new Guid("000C0396-0000-0000-C000-000000000046"); }), "ribbon COM interface");
        Check(!typeof(PowerPointAddIn).Assembly.GetReferencedAssemblies().Any(delegate(AssemblyName name) { return name.Name == "Office" || name.Name == "Microsoft.Office.Interop.PowerPoint"; }), "interop embedded, no PIA deployment dependency");
        Equal(typeof(PowerPointAddIn).GUID.ToString("B").ToUpperInvariant(), InstallerEngine.ClassId, "installer CLSID matches plugin");
    }

    private static void Installer()
    {
        string prefix = @"Software\SlidePace.Tests\" + Guid.NewGuid().ToString("N");
        var options = new InstallOptions
        {
            Root = Path.Combine(root, "install-" + Guid.NewGuid().ToString("N")),
            SettingsRoot = Path.Combine(root, "install-settings-" + Guid.NewGuid().ToString("N")),
            RegistryPrefix = prefix,
            TestMode = true
        };
        var installer = new InstallerEngine(options, null);
        try
        {
            string payload = installer.Install();
            Check(File.Exists(payload), "payload deployed");
            Check(File.Exists(Path.Combine(options.Root, "Uninstall.exe")), "uninstall executable deployed");
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            {
                using (RegistryKey key = registry.OpenSubKey(installer.AddInKey)) Equal((int)key.GetValue("LoadBehavior"), 3, "automatic load registration");
                using (RegistryKey key = registry.OpenSubKey(installer.ClassKey + @"\InprocServer32")) Equal((string)key.GetValue("CodeBase"), new Uri(payload).AbsoluteUri, "COM codebase");
            }
            string updated = installer.Install();
            Check(File.Exists(updated) && !File.Exists(payload), "reinstall switches payload and cleans previous version");
            Equal(Directory.GetDirectories(options.Root).Length, 1, "one deployed version");
            options.FailAfterRegistryForTest = true;
            bool failed = false;
            try { installer.Install(); } catch (IOException) { failed = true; }
            Check(failed, "failure injected");
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            using (RegistryKey key = registry.OpenSubKey(installer.ClassKey + @"\InprocServer32"))
                Equal((string)key.GetValue("CodeBase"), new Uri(updated).AbsoluteUri, "failed update restores original registration");
            Check(File.Exists(updated), "failed update keeps original payload");
            Equal(Directory.GetDirectories(options.Root).Length, 1, "failed update cleans staged files");
            options.FailAfterRegistryForTest = false;
            new SettingsStore(options.SettingsRoot).Save(new UserSettings());
            installer.Uninstall(false);
            Check(File.Exists(Path.Combine(options.SettingsRoot, "settings.json")), "uninstall preserves settings by default");
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                Check(registry.OpenSubKey(installer.AddInKey) == null, "uninstall removes registry");
            Equal(Directory.GetFileSystemEntries(options.Root).Length, 0, "uninstall removes files");
            installer.Install();
            File.WriteAllText(Path.Combine(options.Root, "Uninstall.exe.backup"), "unrelated backup");
            Directory.CreateDirectory(Path.Combine(options.Root, "v-unrelated"));
            File.WriteAllText(Path.Combine(options.Root, "v-unrelated", "notes.txt"), "unrelated content");
            File.WriteAllText(Path.Combine(options.SettingsRoot, "notes.txt"), "unrelated settings content");
            installer.Install();
            Check(File.Exists(Path.Combine(options.Root, "Uninstall.exe.backup")), "unique backup does not overwrite unrelated file");
            installer.Uninstall(true);
            Check(!File.Exists(Path.Combine(options.SettingsRoot, "settings.json")), "optional owned settings removal");
            Check(File.Exists(Path.Combine(options.SettingsRoot, "notes.txt")), "unrelated settings file preserved");
            Check(File.Exists(Path.Combine(options.Root, "v-unrelated", "notes.txt")), "unrelated version directory preserved");
            Check(File.Exists(Path.Combine(options.Root, "Uninstall.exe.backup")), "unrelated installation file preserved");
            options.FailAfterRegistryForTest = true;
            options.Root = Path.Combine(root, "first-install-failure-" + Guid.NewGuid().ToString("N"));
            bool firstFailed = false;
            try { installer.Install(); } catch (IOException) { firstFailed = true; }
            Check(firstFailed && !Directory.Exists(options.Root), "failed first install removes deployment");
            Console.WriteLine("Installer uses isolated test registry prefix; Office registration was not changed.");
        }
        finally
        {
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64)) registry.DeleteSubKeyTree(prefix, false);
        }
    }

    private static void UserInterface()
    {
        Point pointer = Cursor.Position;
        var settings = new UserSettings();
        var time = new FakeTime();
        var engine = new TimerEngine(time);
        int actions = 0;
        using (var manager = new OverlayManager(settings, delegate { }, delegate { actions++; engine.Pause(); }, delegate { engine.Start(); }, delegate { engine.Reset(); }, delegate { }, delegate(TimerMode mode) { engine.SelectMode(mode); }))
        {
            try
            {
                Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Left + 30, Screen.PrimaryScreen.Bounds.Top + 500);
                engine.BeginSlideShow();
                manager.Render(engine.Snapshot(), true);
                Equal(manager.Forms.Count, 0, "none has no windows");
                engine.SetCountdownSeconds(5);
                engine.SelectMode(TimerMode.CountDown);
                manager.Render(engine.Snapshot(), true);
                Pump(200);
                Check(manager.Forms.Count >= 1, "timer windows shown");
                Equal(manager.Forms.Count, Math.Min(2, Screen.AllScreens.Length), "both screens or single shared window");
                Check(manager.Forms.All(delegate(OverlayForm form) { return !form.ControlsVisible; }), "controls initially hidden");
                Check(manager.Forms.All(delegate(OverlayForm form) { return form.Handle != GetForegroundWindow(); }), "floating windows never become foreground windows");
                OverlayForm first = manager.Forms[0];
                Check(first.Opacity > 0 && first.Opacity < 1, "card is translucent");
                Check(first.Height == first.CompactBounds.Height && first.Width < 300, "digit-only card drops the fixed width and hidden button height");
                Check(first.Controls.OfType<Label>().Where(delegate(Label label) { return label.Visible; }).All(delegate(Label label) { return label.Text == first.DisplayText; }), "no mode or status labels are visible");
                CheckVisibleOverlay(first, Path.Combine(root, "overlay-compact.png"));
                using (var dialog = new SettingsForm(settings, false))
                {
                    dialog.Show();
                    ComboBox fontPicker = (ComboBox)dialog.Controls.Find("NumberFont", true).Single();
                    fontPicker.SelectedItem = "Arial";
                    SaveWindow(dialog, Path.Combine(root, "settings-font.png"));
                    Button save = dialog.Controls.OfType<TableLayoutPanel>().Single().Controls.OfType<FlowLayoutPanel>()
                        .SelectMany(delegate(FlowLayoutPanel panel) { return panel.Controls.OfType<Button>(); }).Single(delegate(Button button) { return button.Text == "保存设置"; });
                    save.PerformClick();
                    Equal(dialog.Result.NumberFontName, "Arial", "font picker saves selected family");
                    settings.NumberFontName = dialog.Result.NumberFontName;
                }
                manager.ApplySettings(settings);
                manager.Render(engine.Snapshot(), true);
                first = manager.Forms[0];
                Check(manager.Forms.All(delegate(OverlayForm form) { return form.DisplayFontName == "Arial"; }), "font selection applies to both screens");
                Rectangle compactBefore = first.CompactBounds;
                Cursor.Position = new Point(compactBefore.Left + compactBefore.Width / 2, compactBefore.Top + compactBefore.Height / 2);
                manager.Render(engine.Snapshot(), true);
                Check(first.ControlsVisible, "hover digits reveals controls");
                Check(first.Height > compactBefore.Height && first.CompactBounds == compactBefore, "hover expands controls without moving digits");
                if (manager.Forms.Count == 2) Check(!manager.Forms[1].ControlsVisible, "hover is local to one screen");
                var buttons = first.Controls.OfType<Panel>().Single().Controls.OfType<Button>().ToList();
                Button pauseButton = buttons.Single(delegate(Button button) { return button.Text == "暂停"; });
                Cursor.Position = pauseButton.PointToScreen(new Point(pauseButton.Width / 2, pauseButton.Height / 2));
                manager.Render(engine.Snapshot(), true);
                Check(first.ControlsVisible, "moving into button row keeps controls visible");
                buttons.Single(delegate(Button button) { return button.Text == "暂停"; }).PerformClick();
                Equal(actions, 1, "actual button dispatches control action");
                Check(!engine.Snapshot().IsRunning, "pause button pauses timer");
                engine.Start();
                time.Seconds = 7;
                manager.Render(engine.Snapshot(), true);
                Check(manager.Forms.All(delegate(OverlayForm form) { return form.DisplayText == "-00:00:02" && form.DisplayColor.ToArgb() == Color.Red.ToArgb(); }), "overtime red and same readout on both screens");
                settings.OvertimeColor = "#22CC88";
                manager.Render(engine.Snapshot(), true);
                Check(first.DisplayColor.ToArgb() == Color.FromArgb(0x22, 0xCC, 0x88).ToArgb(), "custom overtime color");
                SaveWindow(first, Path.Combine(root, "overlay-overtime.png"));
                Cursor.Position = new Point(Screen.PrimaryScreen.Bounds.Left + 30, Screen.PrimaryScreen.Bounds.Top + 500);
                manager.Render(engine.Snapshot(), true);
                Check(!first.ControlsVisible && first.Height == compactBefore.Height && first.Right == compactBefore.Right && first.Top == compactBefore.Top,
                    "leaving hover removes button height, overtime width keeps right alignment");
                CheckVisibleOverlay(first, Path.Combine(root, "overlay-overtime-compact.png"));
                settings.PresenterPosition = settings.AudiencePosition = OverlayPosition.BottomRight;
                manager.ApplySettings(settings);
                manager.Render(engine.Snapshot(), true);
                first = manager.Forms[0];
                Rectangle bottomCompact = first.CompactBounds;
                Cursor.Position = new Point(bottomCompact.Left + bottomCompact.Width / 2, bottomCompact.Top + bottomCompact.Height / 2);
                manager.Render(engine.Snapshot(), true);
                Check(first.ControlsVisible && Screen.FromControl(first).Bounds.Contains(first.Bounds), "bottom hover expands inside screen bounds");
                Equal(first.CompactBounds, bottomCompact, "bottom hover preserves digit position while expanding upward");
                engine.Reset();
                manager.Render(engine.Snapshot(), true);
                Equal(first.DisplayColor.ToArgb(), Color.White.ToArgb(), "reset restores normal color");
                engine.SelectMode(TimerMode.Clock);
                manager.Render(engine.Snapshot(), true);
                Check(!first.ControlsVisible, "clock has no hover controls");
                engine.SelectMode(TimerMode.None);
                manager.Render(engine.Snapshot(), true);
                Equal(manager.Forms.Count, 0, "none closes all windows");
            }
            finally { Cursor.Position = pointer; }
        }
    }

    private static void OfficeIntegration()
    {
        PowerPoint.Application app = null;
        PowerPoint.Presentation presentation = null;
        PowerPoint.DocumentWindow previous = null;
        bool ownApplication = false;
        var plugin = new PowerPointAddIn(Path.Combine(root, "office-config-" + Guid.NewGuid().ToString("N")));
        Point pointer = Cursor.Position;
        Array custom = new object[0];
        try
        {
            try { app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
            catch (COMException) { app = (PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")); ownApplication = true; }
            int originalCount = app.Presentations.Count;
            try { previous = app.ActiveWindow; } catch (COMException) { }
            if (app.SlideShowWindows.Count > 0) throw new InvalidOperationException("An existing user slide show is active; integration test will not disturb it.");
            plugin.OnConnection(app, 0, null, ref custom);
            Check(plugin.Engine != null, "connection initialized in actual PowerPoint");
            presentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
            var slide = presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(18, 32, 51));
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 60, 150, 600, 100);
            title.TextFrame.TextRange.Text = "SlidePace · PowerPoint 集成验证";
            title.TextFrame.TextRange.Font.Size = 32;
            title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            presentation.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutBlank);
            presentation.SlideShowSettings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeWindow;
            PowerPoint.SlideShowWindow show = presentation.SlideShowSettings.Run();
            show.Activate();
            ActivateTestWindow(new IntPtr(show.HWND));
            Console.WriteLine("Windowed host: " + ShowWindows.Describe(new IntPtr(show.HWND)));
            Console.WriteLine("Windowed root: " + ShowWindows.Describe(ShowWindows.Root(new IntPtr(show.HWND))));
            Console.WriteLine("Host editor: " + ShowWindows.Describe(new IntPtr(app.HWND)));
            Console.WriteLine("Foreground: " + ShowWindows.Describe(GetForegroundWindow()));
            Until(delegate { return plugin.Engine.IsSlideShowActive; }, 5000, "actual show auto detection");
            Check(plugin.BeginEventCount > 0, "PowerPoint raised SlideShowBegin, not just polling");
            Equal(plugin.Overlays.Forms.Count, 0, "none in actual show displays no box");
            plugin.SelectMode(TimerMode.CountUp);
            Pump(1200);
            Check(plugin.Engine.Snapshot().IsRunning && plugin.Engine.Snapshot().ElapsedSeconds >= 1, "actual host auto starts count up");
            Equal(plugin.Overlays.Forms.Count, 1, "windowed show has one overlay inside the actual show");
            System.Drawing.Rectangle showArea;
            Check(ShowWindows.TryGetBounds(new IntPtr(show.HWND), out showArea) && showArea.Contains(plugin.Overlays.Forms[0].Bounds), "actual overlay stays in show client area");
            Check(plugin.Overlays.Forms[0].TopMost && GetWindow(plugin.Overlays.Forms[0].Handle, 4) == ShowWindows.Root(new IntPtr(show.HWND)),
                "actual overlay stays above and is owned by the show root");
            Check(plugin.Overlays.Forms.Select(delegate(OverlayForm form) { return form.DisplayText; }).Distinct().Count() == 1, "actual displays synchronized");
            show.View.Next();
            Pump(300);
            Check(plugin.Engine.Snapshot().ElapsedSeconds >= 1, "turning slides does not reset timer");
            plugin.StartPause();
            double paused = plugin.Engine.Snapshot().ElapsedSeconds;
            Pump(250);
            Check(Math.Abs(plugin.Engine.Snapshot().ElapsedSeconds - paused) < 0.001, "actual host pause");
            plugin.Continue();
            Pump(200);
            Check(plugin.Engine.Snapshot().ElapsedSeconds > paused, "actual host resume");
            plugin.Reset();
            Pump(200);
            Check(!plugin.Engine.Snapshot().IsRunning && plugin.Engine.Snapshot().Text == "00:00:00", "reset remains stopped");
            var settings = new UserSettings { Mode = (int)TimerMode.CountDown, CountdownSeconds = 2, OvertimeColor = "#22CC88" };
            plugin.ApplySettings(settings);
            Pump(3300);
            Check(plugin.Engine.Snapshot().IsRunning && plugin.Engine.Snapshot().IsOvertime, "actual countdown continues past zero");
            Check(plugin.Overlays.Forms.All(delegate(OverlayForm form) { return form.DisplayText.StartsWith("-") && form.DisplayColor.ToArgb() == Color.FromArgb(0x22, 0xCC, 0x88).ToArgb(); }), "actual overtime color on both screens");
            Rectangle compact = plugin.Overlays.Forms[0].CompactBounds;
            Cursor.Position = new Point(compact.Left + compact.Width / 2, compact.Top + compact.Height / 2);
            Pump(150);
            CheckVisibleOverlay(plugin.Overlays.Forms[0], Path.Combine(root, "powerpoint-overlay.png"));
            plugin.SelectMode(TimerMode.Clock);
            Pump(150);
            Check(plugin.Overlays.Forms.All(delegate(OverlayForm form) { return !form.ControlsVisible; }), "actual clock hides controls");
            plugin.SelectMode(TimerMode.None);
            Equal(plugin.Overlays.Forms.Count, 0, "actual cancel clears both screens");
            plugin.SelectMode(TimerMode.CountUp);
            // Reproduce switching back to editing while a show object still exists.
            Console.WriteLine("Office stage: switch to editor");
            presentation.Windows[1].Activate();
            ActivateTestWindow(new IntPtr(app.HWND));
            Pump(700);
            Equal(plugin.Overlays.Forms.Count, 0, "returning to editor hides the running-show overlay");
            Console.WriteLine("Office stage: return to windowed show");
            show.Activate();
            ActivateTestWindow(new IntPtr(show.HWND));
            Pump(700);
            Check(plugin.Overlays.Forms.Count == 1, "returning to the show restores its overlay");
            show.View.Exit();
            Until(delegate { return !plugin.Engine.IsSlideShowActive; }, 5000, "actual end auto pauses");
            Check(plugin.EndEventCount > 0, "PowerPoint raised SlideShowEnd");
            Equal(plugin.Overlays.Forms.Count, 0, "actual end removes windows");
            plugin.SelectMode(TimerMode.Clock);
            Pump(200);
            Equal(plugin.Overlays.Forms.Count, 0, "clock selected in editor does not create overlay");
            plugin.SelectMode(TimerMode.CountUp);
            Console.WriteLine("Office stage: repeat windowed show");
            show = presentation.SlideShowSettings.Run();
            Until(delegate { return plugin.Engine.IsSlideShowActive; }, 5000, "repeat show detected");
            Check(plugin.Engine.Snapshot().IsRunning, "repeat show auto resumes");
            show.View.Exit();
            Pump(200);
            presentation.SlideShowSettings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeSpeaker;
            presentation.SlideShowSettings.ShowPresenterView = Screen.AllScreens.Length > 1 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            Console.WriteLine("Office stage: start full-screen presenter show");
            show = presentation.SlideShowSettings.Run();
            show.Activate();
            ActivateTestWindow(new IntPtr(show.HWND));
            Pump(1200);
            Console.WriteLine("Full-screen host: " + ShowWindows.Describe(new IntPtr(show.HWND)));
            Equal(plugin.Overlays.Forms.Count, Math.Min(2, Screen.AllScreens.Length), "full-screen audience and actual presenter view both have overlays");
            foreach (OverlayForm form in plugin.Overlays.Forms)
            {
                Check(form.AreaBounds.Contains(form.Bounds), "full-screen overlay stays inside its view");
                CheckVisibleOverlay(form, Path.Combine(root, form.IsPresenter ? "presenter-overlay.png" : "audience-overlay.png"));
            }
            show.View.Exit();
            Pump(300);
            plugin.OnDisconnection(0, ref custom);
            presentation.Saved = Office.MsoTriState.msoTrue;
            presentation.Close();
            presentation = null;
            Equal(app.Presentations.Count, originalCount, "user presentations preserved");
            Console.WriteLine("Actual Office version: " + app.Version + "; begin events: " + plugin.BeginEventCount + "; end events: " + plugin.EndEventCount);
        }
        finally
        {
            plugin.OnDisconnection(0, ref custom);
            if (presentation != null)
            {
                try { if (app.SlideShowWindows.Count > 0) foreach (PowerPoint.SlideShowWindow show in app.SlideShowWindows) if (show.Presentation == presentation) show.View.Exit(); } catch { }
                try { presentation.Saved = Office.MsoTriState.msoTrue; presentation.Close(); } catch { }
            }
            try { if (previous != null) previous.Activate(); } catch { }
            if (ownApplication && app != null) try { app.Quit(); } catch { }
            Cursor.Position = pointer;
        }
    }
    private static void SaveWindow(Form form, string path)
    {
        Pump(50);
        using (var image = new Bitmap(form.Width, form.Height))
        using (Graphics graphics = Graphics.FromImage(image))
        {
            graphics.CopyFromScreen(form.Location, Point.Empty, form.Size);
            image.Save(path, ImageFormat.Png);
        }
    }
    private static void CheckVisibleOverlay(OverlayForm form, string path)
    {
        Pump(100);
        using (var image = new Bitmap(form.Width, form.Height))
        using (Graphics graphics = Graphics.FromImage(image))
        {
            graphics.CopyFromScreen(form.Location, Point.Empty, form.Size);
            image.Save(path, ImageFormat.Png);
            // Read desktop pixels, not DrawToBitmap: a covered HWND must fail.
            Label digits = form.Controls.OfType<Label>().Single(delegate(Label label) { return label.AccessibleName == "时间"; });
            Color foreground = form.DisplayColor;
            double opacity = form.Opacity;
            int matching = 0;
            int left = image.Width, top = image.Height, right = -1, bottom = -1;
            for (int y = digits.Top; y < digits.Bottom; y++)
                for (int x = digits.Left; x < digits.Right; x++)
                {
                    Color pixel = image.GetPixel(x, y);
                    if (new[] { pixel.R, pixel.G, pixel.B }.Zip(new[] { foreground.R, foreground.G, foreground.B },
                        delegate(byte actual, byte expected) { return actual >= expected * opacity - 3 && actual <= expected * opacity + 255 * (1 - opacity) + 3; })
                        .All(delegate(bool valid) { return valid; }))
                    {
                        matching++;
                        left = Math.Min(left, x); right = Math.Max(right, x);
                        top = Math.Min(top, y); bottom = Math.Max(bottom, y);
                    }
                }
            Check(matching >= 50, "screen pixels contain visible time digits: " + Path.GetFileName(path) + " (" + matching + " pixels)");
            int compactTop = form.CompactBounds.Top - form.Top;
            int[] margins = { left, image.Width - 1 - right, top - compactTop, compactTop + form.CompactBounds.Height - 1 - bottom };
            Check(margins.Max() - margins.Min() <= 3, "visible digit margins match on all sides: " + string.Join(",", margins));
        }
    }
    [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool AttachThreadInput(uint first, uint second, bool attach);
    [DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();
    private static void ActivateTestWindow(IntPtr window)
    {
        IntPtr rootWindow = ShowWindows.Root(window);
        uint process;
        uint foregroundThread = GetWindowThreadProcessId(GetForegroundWindow(), out process);
        uint targetThread = GetWindowThreadProcessId(rootWindow, out process);
        uint current = GetCurrentThreadId();
        bool foregroundAttached = foregroundThread != 0 && foregroundThread != current && AttachThreadInput(current, foregroundThread, true);
        bool targetAttached = targetThread != 0 && targetThread != current && targetThread != foregroundThread && AttachThreadInput(current, targetThread, true);
        try { SetForegroundWindow(rootWindow); }
        finally
        {
            if (targetAttached) AttachThreadInput(current, targetThread, false);
            if (foregroundAttached) AttachThreadInput(current, foregroundThread, false);
        }
    }
}
