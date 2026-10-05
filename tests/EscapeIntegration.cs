using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using System.Windows.Forms;
using SlidePace;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static partial class Program
{
    [StructLayout(LayoutKind.Sequential)]
    private struct InputMouse
    {
        public int X, Y;
        public uint Data, Flags, Time;
        public IntPtr Extra;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct InputKeyboard
    {
        public ushort Key, Scan;
        public uint Flags, Time;
        public IntPtr Extra;
    }
    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public InputMouse Mouse;
        [FieldOffset(0)] public InputKeyboard Keyboard;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct NativeInput { public uint Type; public InputUnion Value; }
    [StructLayout(LayoutKind.Sequential)]
    private struct GuiThreadInfo
    {
        public uint Size, Flags;
        public IntPtr Active, Focus, Capture, MenuOwner, MoveSize, Caret;
        public WindowRectangle CaretRectangle;
    }
    [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, NativeInput[] inputs, int size);
    [DllImport("user32.dll")] private static extern bool GetGUIThreadInfo(uint thread, ref GuiThreadInfo info);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
    [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr window, int index);

    private static void KeyInput(ushort key)
    {
        var inputs = new[]
        {
            new NativeInput { Type = 1, Value = new InputUnion { Keyboard = new InputKeyboard { Key = key } } },
            new NativeInput { Type = 1, Value = new InputUnion { Keyboard = new InputKeyboard { Key = key, Flags = 2 } } }
        };
        Check(SendInput(2, inputs, Marshal.SizeOf(typeof(NativeInput))) == 2, "physical keyboard input delivered");
    }
    private static void ClickInput(Point point)
    {
        Cursor.Position = point;
        var inputs = new[]
        {
            new NativeInput { Value = new InputUnion { Mouse = new InputMouse { Flags = 2 } } },
            new NativeInput { Value = new InputUnion { Mouse = new InputMouse { Flags = 4 } } }
        };
        Check(SendInput(2, inputs, Marshal.SizeOf(typeof(NativeInput))) == 2, "physical mouse click delivered");
    }
    private static string WindowClass(IntPtr window)
    {
        var value = new StringBuilder(256);
        GetClassName(window, value, value.Capacity);
        return window + ":" + value;
    }
    private static List<IntPtr> PresenterWindows(uint process)
    {
        var result = new List<IntPtr>();
        EnumWindows(delegate(IntPtr window, IntPtr parameter)
        {
            uint candidate;
            GetWindowThreadProcessId(window, out candidate);
            if (candidate == process && IsWindowVisible(window))
            {
                var title = new StringBuilder(256);
                GetWindowText(window, title, title.Capacity);
                if (WindowClass(window).IndexOf("Presenter", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    title.ToString().IndexOf("演示者视图", StringComparison.Ordinal) >= 0 ||
                    title.ToString().IndexOf("Presenter View", StringComparison.OrdinalIgnoreCase) >= 0)
                    result.Add(window);
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }
    private static List<IntPtr> NativeTimerWindows(uint process)
    {
        var result = new List<IntPtr>();
        EnumWindows(delegate(IntPtr window, IntPtr parameter)
        {
            uint candidate;
            GetWindowThreadProcessId(window, out candidate);
            var title = new StringBuilder(256);
            if (candidate == process && IsWindowVisible(window))
            {
                GetWindowText(window, title, title.Capacity);
                if (title.ToString() == "SlidePace 计时") result.Add(window);
            }
            return true;
        }, IntPtr.Zero);
        return result;
    }
    private static void TraceEscape(string stage, PowerPoint.Application app, object automation, uint process)
    {
        IntPtr foreground = GetForegroundWindow();
        uint pid;
        uint thread = GetWindowThreadProcessId(foreground, out pid);
        var gui = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
        GetGUIThreadInfo(thread, ref gui);
        Console.WriteLine(DateTime.Now.ToString("HH:mm:ss.fff") + " " + stage + " | shows=" + app.SlideShowWindows.Count + " | status=" + Status(automation) +
            " | foreground=" + WindowClass(foreground) + " pid=" + pid + " tid=" + thread +
            " | active=" + WindowClass(gui.Active) + " focus=" + WindowClass(gui.Focus) + " capture=" + WindowClass(gui.Capture) +
            " menu=" + WindowClass(gui.MenuOwner) + " flags=" + gui.Flags);
        foreach (IntPtr presenter in PresenterWindows(process))
            Console.WriteLine("  presenter " + ShowWindows.Describe(presenter) + " owner=" + WindowClass(GetWindow(presenter, 4)));
        for (int index = 1; index <= app.SlideShowWindows.Count; index++)
        {
            var show = app.SlideShowWindows[index];
            Console.WriteLine("  audience " + ShowWindows.Describe(new IntPtr(show.HWND)) + " state=" + show.View.State);
        }
        foreach (IntPtr timer in NativeTimerWindows(process))
            Console.WriteLine("  timer " + ShowWindows.Describe(timer) + " owner=" + WindowClass(GetWindow(timer, 4)));
    }
    private static void SelectInstalledMode(AutomationElement host, object automation, TimerMode target)
    {
        TimerMode current = (TimerMode)Enum.Parse(typeof(TimerMode), Status(automation).Split('|')[0]);
        if (current == target) return;
        AutomationElement tab = host.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "SlidePace 计时"));
        Check(tab != null, "installed ribbon tab found");
        SelectAutomationElement(tab);
        Pump(200);
        TimerMode buttonMode = target == TimerMode.None ? current : target;
        string name = buttonMode == TimerMode.CountUp ? "顺计时" : buttonMode == TimerMode.CountDown ? "倒计时" : "系统时间";
        var button = host.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, name));
        Check(button != null, "installed mode button found: " + name);
        ToggleAutomationElement(button);
        Pump(200);
        Equal(Status(automation).Split('|')[0], target.ToString(), "installed ribbon mode selection");
    }
    private static void RequireHostForeground(uint process)
    {
        uint current;
        GetWindowThreadProcessId(GetForegroundWindow(), out current);
        Check(current == process, "keyboard destination belongs to PowerPoint; no input sent to another application");
    }
    private static void OverlayFocusIntegration()
    {
        Point originalPointer = Cursor.Position;
        try
        {
        using (var owner = new Form { Text = "SlidePace temporary keyboard owner", Bounds = new Rectangle(100, 100, 800, 500) })
        using (var input = new TextBox { Location = new Point(20, 20) })
        using (var manager = new OverlayManager(new UserSettings(), delegate { }, delegate { }, delegate { }, delegate { }, delegate { }, delegate(TimerMode mode) { }))
        {
            owner.Controls.Add(input);
            owner.Show();
            ActivateTestWindow(owner.Handle);
            input.Focus();
            Cursor.Position = new Point(owner.Left + 40, owner.Bottom - 50);
            Pump(100);
            Console.WriteLine("UI before render foreground=" + WindowClass(GetForegroundWindow()) + " owner=" + WindowClass(owner.Handle));
            IntPtr focus = input.Handle;
            manager.SetHostWindows(owner.Handle, owner.Handle);
            var time = new FakeTime();
            var engine = new TimerEngine(time);
            engine.SelectMode(TimerMode.CountDown);
            engine.BeginSlideShow();
            manager.Render(engine.Snapshot(), true);
            Console.WriteLine("UI after render foreground=" + WindowClass(GetForegroundWindow()) + " timer=" + WindowClass(manager.Forms.Single().Handle));
            Pump(200);
            Equal(GetForegroundWindow(), owner.Handle, "first display leaves foreground with host");
            var gui = new GuiThreadInfo { Size = (uint)Marshal.SizeOf(typeof(GuiThreadInfo)) };
            GetGUIThreadInfo(GetCurrentThreadId(), ref gui);
            Equal(gui.Focus, focus, "first display retains host child keyboard focus");
            for (int index = 0; index < 5; index++)
            {
                time.Seconds++;
                manager.Render(engine.Snapshot(), true);
                Pump(100);
            }
            Equal(GetForegroundWindow(), owner.Handle, "refresh does not activate timer");
            var timer = manager.Forms.Single();
            Check((GetWindowLong(timer.Handle, -20) & 8) != 0, "timer is natively topmost without WinForms activation");
            Cursor.Position = new Point(timer.Left + timer.Width / 2, timer.Top + 20);
            manager.Render(engine.Snapshot(), true);
            Pump(100);
            Check(timer.ControlsVisible, "hover expands controls");
            Equal(GetForegroundWindow(), owner.Handle, "hover does not activate timer");
            GetGUIThreadInfo(GetCurrentThreadId(), ref gui);
            Equal(gui.Focus, focus, "hover retains host child focus");
            Cursor.Position = new Point(owner.Left + 30, owner.Top + 100);
            manager.ApplySettings(new UserSettings { AudiencePosition = OverlayPosition.BottomLeft });
            manager.Render(engine.Snapshot(), true);
            Pump(100);
            Equal(GetForegroundWindow(), owner.Handle, "recreated timer does not activate");
            GetGUIThreadInfo(GetCurrentThreadId(), ref gui);
            Equal(gui.Focus, focus, "recreated timer retains host child focus");
            manager.CloseForms();
            Equal(GetForegroundWindow(), owner.Handle, "timer cleanup retains foreground");
        }
        }
        finally { Cursor.Position = originalPointer; }
    }
    private static string[] RegisterEscapeHost()
    {
        const string clsid = "{86B4DB30-2296-4B59-938D-11A06F843410}";
        const string progid = "SlidePace.EscapeValidation";
        string[] keys = { "Software/Classes/CLSID/" + clsid, "Software/Classes/" + progid,
            "Software/Microsoft/Office/PowerPoint/Addins/" + progid };
        keys = keys.Select(delegate(string key) { return key.Replace('/', '\\'); }).ToArray();
        using (var registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
        {
            foreach (string path in keys) using (var existing = registry.OpenSubKey(path)) Check(existing == null, "temporary COM identity is unused");
            try
            {
                using (var key = registry.CreateSubKey(keys[0] + "\\InprocServer32"))
                {
                    key.SetValue("", "mscoree.dll");
                    key.SetValue("ThreadingModel", "Both");
                    key.SetValue("Assembly", typeof(NativeEscapeHost).Assembly.FullName);
                    key.SetValue("Class", typeof(NativeEscapeHost).FullName);
                    key.SetValue("RuntimeVersion", "v4.0.30319");
                    key.SetValue("CodeBase", new Uri(typeof(NativeEscapeHost).Assembly.Location).AbsoluteUri);
                }
                using (var key = registry.CreateSubKey(keys[0] + "\\ProgId")) key.SetValue("", progid);
                using (var key = registry.CreateSubKey(keys[1] + "\\CLSID")) key.SetValue("", clsid);
                using (var key = registry.CreateSubKey(keys[2]))
                {
                    key.SetValue("FriendlyName", "SlidePace temporary Esc verification");
                    key.SetValue("Description", "Workspace-only verification; removed after test");
                    key.SetValue("LoadBehavior", 0, RegistryValueKind.DWord);
                }
            }
            catch
            {
                foreach (string path in keys) registry.DeleteSubKeyTree(path, false);
                throw;
            }
        }
        return keys;
    }
    private static void SelectEscapeMode(AutomationElement host, object automation, TimerMode mode, bool temporaryNative)
    {
        if (!temporaryNative) { SelectInstalledMode(host, automation, mode); return; }
        automation.GetType().InvokeMember("SetModeForValidation", BindingFlags.InvokeMethod, null, automation, new object[] { (int)mode });
        Equal(Status(automation).Split('|')[0], mode.ToString(), "native validation mode selected");
    }
    private static void CleanupEscape(Action action, string description, List<string> errors)
    {
        try { action(); }
        catch (Exception error) { errors.Add(description + ": " + error.Message); }
    }
    private static void SetPresenterForTest(object automation, bool visible)
    {
        automation.GetType().InvokeMember("SetPresenterTimerForValidation", BindingFlags.InvokeMethod, null, automation, new object[] { visible });
    }
    private static double NativeElapsed(object automation)
    {
        return (double)automation.GetType().InvokeMember("GetElapsedForValidation", BindingFlags.InvokeMethod, null, automation, new object[0]);
    }
    private static void CheckAudienceOnly(PowerPoint.Application app, object automation, uint process, TimerMode mode, IList<IntPtr> presenters, string name)
    {
        Equal(app.SlideShowWindows.Count, 1, "hiding presenter timer keeps audience slideshow running");
        Check(PresenterWindows(process).SequenceEqual(presenters), "hiding timer preserves actual presenter view HWND");
        Check(app.SlideShowWindows[1].View.State == PowerPoint.PpSlideShowState.ppSlideShowRunning, "hiding timer does not pause PowerPoint show");
        List<IntPtr> timers = NativeTimerWindows(process);
        Equal(timers.Count, mode == TimerMode.None ? 0 : 1, "only audience timer remains when presenter timer is hidden");
        Equal(Status(automation).Split('|')[3], timers.Count.ToString(), "managed form count matches native visible windows");
        Check(!timers.Contains(GetForegroundWindow()), "visibility change does not steal keyboard focus");
        foreach (IntPtr presenter in presenters)
            SaveHostWindow(presenter, Path.Combine(root, name + "-presenter-without-timer.png"));
        if (mode == TimerMode.None) return;
        IntPtr timer = timers.Single();
        Equal(GetWindow(timer, 4), ShowWindows.Root(new IntPtr(app.SlideShowWindows[1].HWND)), "remaining timer belongs to audience window");
        WindowRectangle area;
        Check(GetWindowRect(timer, out area), "audience timer rectangle available");
        using (var bitmap = new Bitmap(area.Right - area.Left, area.Bottom - area.Top))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.CopyFromScreen(new Point(area.Left, area.Top), Point.Empty, bitmap.Size);
            int white = 0;
            for (int y = 0; y < bitmap.Height; y++)
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    if (pixel.R >= 200 && pixel.G >= 200 && pixel.B >= 200) white++;
                }
            Check(white >= 50, "audience time digits remain visible above real slideshow");
            bitmap.Save(Path.Combine(root, name + "-audience-only.png"));
        }
        SaveHostWindow(ShowWindows.Root(new IntPtr(app.SlideShowWindows[1].HWND)), Path.Combine(root, name + "-audience-view.png"));
    }
    private static void EscapeIntegration(bool temporaryNative, bool testPresenterVisibility = false)
    {
        PowerPoint.Application app;
        bool ownApplication = false;
        try { app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
        catch (COMException) { app = (PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")); ownApplication = true; }
        Check(app.SlideShowWindows.Count == 0, "no existing user slideshow will be disturbed");
        object id = "SlidePace.PowerPointAddIn";
        var addin = app.COMAddIns.Item(ref id);
        Check(addin.Connect && addin.Object != null, "installed native plugin is connected");
        object installedAutomation = addin.Object;
        object automation = installedAutomation;
        TimerMode originalMode = (TimerMode)Enum.Parse(typeof(TimerMode), Status(automation).Split('|')[0]);
        PowerPoint.DocumentWindow previous = null;
        try { previous = app.ActiveWindow; } catch { }
        Point pointer = Cursor.Position;
        int originalCount = app.Presentations.Count;
        int finalCount = -1;
        uint process;
        GetWindowThreadProcessId(new IntPtr(app.HWND), out process);
        PowerPoint.Presentation presentation = null;
        Office.COMAddIn validationAddin = null;
        string[] temporaryKeys = null;
        var failures = new List<string>();
        var cleanupErrors = new List<string>();
        try
        {
            presentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
            for (int index = 1; index <= 2; index++)
            {
                var slide = presentation.Slides.Add(index, PowerPoint.PpSlideLayout.ppLayoutBlank);
                slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
                slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(18, 32, 51));
                var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 80, 180, 650, 100);
                title.TextFrame.TextRange.Text = "SlidePace · Esc 验证（临时文稿）";
                title.TextFrame.TextRange.Font.Size = 32;
                title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            }
            presentation.SlideShowSettings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeSpeaker;
            presentation.SlideShowSettings.ShowPresenterView = Screen.AllScreens.Length > 1 ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            var host = AutomationElement.FromHandle(new IntPtr(app.HWND));
            if (temporaryNative)
            {
                SelectInstalledMode(host, installedAutomation, TimerMode.None);
                temporaryKeys = RegisterEscapeHost();
                app.COMAddIns.Update();
                object validationId = "SlidePace.EscapeValidation";
                validationAddin = app.COMAddIns.Item(ref validationId);
                validationAddin.Connect = true;
                Until(delegate { return validationAddin.Object != null; }, 5000, "Office loads temporary native test host");
                automation = validationAddin.Object;
                string assembly = (string)automation.GetType().InvokeMember("GetDiagnosticAssembly", BindingFlags.InvokeMethod, null, automation, new object[0]);
                Console.WriteLine("Native production assembly: " + assembly);
                Check(assembly.Split('|')[0] == typeof(PowerPointAddIn).Assembly.FullName, "Office loaded current production assembly identity");
                Check(Path.GetFullPath(assembly.Split('|')[1]) == Path.GetFullPath(typeof(PowerPointAddIn).Assembly.Location), "Office loaded exact built production DLL");
            }
            TimerMode[] modes = temporaryNative ? new[] { TimerMode.CountDown, TimerMode.None, TimerMode.CountUp, TimerMode.Clock } : new[] { TimerMode.CountDown, TimerMode.None };
            foreach (TimerMode mode in modes)
            {
                SelectEscapeMode(host, automation, mode, temporaryNative);
                for (int scenario = 0; scenario < (testPresenterVisibility ? 3 : 2); scenario++)
                {
                    string name = mode + (scenario == 0 ? "-untouched" : scenario == 1 ? "-presenter-click" : "-presenter-hidden");
                    if (testPresenterVisibility) SetPresenterForTest(automation, scenario != 2);
                    presentation.Windows[1].Activate();
                    // Setup only: give the temporary editor input focus BEFORE F5.
                    // No Activate/SetForegroundWindow is permitted after slideshow startup.
                    ActivateTestWindow(new IntPtr(app.HWND));
                    Cursor.Position = new Point(40, 1000);
                    RequireHostForeground(process);
                    TraceEscape(name + " before F5", app, automation, process);
                    KeyInput(0x74);
                    Until(delegate { return app.SlideShowWindows.Count == 1; }, 5000, "F5 starts temporary slideshow");
                    Pump(1600);
                    TraceEscape(name + " startup", app, automation, process);
                    if (temporaryNative)
                    {
                        Check(app.SlideShowWindows[1].View.State == PowerPoint.PpSlideShowState.ppSlideShowRunning, "timer display does not pause PowerPoint slideshow");
                        Check(!NativeTimerWindows(process).Contains(GetForegroundWindow()), "startup focus does not belong to timer");
                        Equal(int.Parse(Status(automation).Split('|')[3]), mode == TimerMode.None ? 0 : scenario == 2 ? 1 : Math.Min(2, Screen.AllScreens.Length), "native overlay count at untouched startup");
                    }
                    if (testPresenterVisibility && scenario == 1)
                    {
                        var presenters = PresenterWindows(process);
                        double elapsed = NativeElapsed(automation);
                        SetPresenterForTest(automation, false);
                        Pump(350);
                        TraceEscape(name + " presenter timer hidden", app, automation, process);
                        Console.WriteLine("  managed overlays: " + automation.GetType().InvokeMember("GetOverlayDetailsForValidation", BindingFlags.InvokeMethod, null, automation, new object[0]));
                        CheckAudienceOnly(app, automation, process, mode, presenters, name);
                        if (mode == TimerMode.CountDown || mode == TimerMode.CountUp)
                            Check(Status(automation).Split('|')[2] == "True" && NativeElapsed(automation) > elapsed, "shared counter continues without pause or reset while presenter timer is hidden");
                        SetPresenterForTest(automation, true);
                        Pump(250);
                        Equal(int.Parse(Status(automation).Split('|')[3]), mode == TimerMode.None ? 0 : Math.Min(2, Screen.AllScreens.Length), "re-enabling restores presenter timer without enabling deselected modes");
                        Equal(NativeTimerWindows(process).Count, int.Parse(Status(automation).Split('|')[3]), "restored timer forms are actually visible");
                        Check(PresenterWindows(process).SequenceEqual(presenters), "re-enabling timer leaves presenter view unchanged");
                    }
                    if (testPresenterVisibility && scenario == 2)
                        CheckAudienceOnly(app, automation, process, mode, PresenterWindows(process), name);
                    if (scenario != 0)
                    {
                        var presenters = PresenterWindows(process);
                        Check(presenters.Count == 1, "actual presenter view available for click");
                        WindowRectangle area;
                        GetWindowRect(presenters[0], out area);
                        ClickInput(new Point(area.Left + (area.Right - area.Left) / 3, area.Top + (area.Bottom - area.Top) / 2));
                        Pump(250);
                        TraceEscape(name + " clicked", app, automation, process);
                    }
                    int endsBefore = int.Parse(Status(automation).Split('|')[5]);
                    RequireHostForeground(process);
                    KeyInput(0x1B);
                    Pump(1200);
                    TraceEscape(name + " after Esc", app, automation, process);
                    if (temporaryNative)
                    {
                        Check(app.SlideShowWindows.Count == 0, "actual Esc exits both audience and presenter show: " + name);
                        Equal(PresenterWindows(process).Count, 0, "presenter view closes on Esc");
                        Equal(int.Parse(Status(automation).Split('|')[5]), endsBefore + 1, "actual SlideShowEnd event occurs once");
                        Equal(Status(automation).Split('|')[2], "False", "Esc pauses timer");
                        Equal(Status(automation).Split('|')[3], "0", "Esc removes all timer forms");
                    }
                    if (app.SlideShowWindows.Count != 0)
                    {
                        failures.Add(name + ": Esc did not end slideshow");
                        foreach (IntPtr presenter in PresenterWindows(process)) SaveHostWindow(presenter, Path.Combine(root, name + "-after-esc.png"));
                        var show = app.SlideShowWindows[1];
                        WindowRectangle area;
                        GetWindowRect(ShowWindows.Root(new IntPtr(show.HWND)), out area);
                        ClickInput(new Point(area.Left + (area.Right - area.Left) / 2, area.Top + (area.Bottom - area.Top) / 2));
                        Pump(250);
                        TraceEscape(name + " audience clicked", app, automation, process);
                        RequireHostForeground(process);
                        KeyInput(0x1B);
                        Pump(1000);
                        TraceEscape(name + " audience Esc", app, automation, process);
                        Check(app.SlideShowWindows.Count == 0, "click audience then Esc ends temporary show");
                    }
                    if (Status(automation).Split('|')[2] != "False" || NativeTimerWindows(process).Count != 0)
                        failures.Add(name + ": timer not cleaned up after exit");
                }
            }
        }
        finally
        {
            if (presentation != null)
            {
                CleanupEscape(delegate { foreach (PowerPoint.SlideShowWindow show in app.SlideShowWindows) if (show.Presentation == presentation) show.View.Exit(); }, "exit temporary slideshow", cleanupErrors);
                CleanupEscape(delegate { presentation.Saved = Office.MsoTriState.msoTrue; presentation.Close(); }, "close temporary presentation", cleanupErrors);
            }
            if (validationAddin != null) CleanupEscape(delegate { validationAddin.Connect = false; }, "disconnect validation add-in", cleanupErrors);
            if (temporaryKeys != null)
            {
                foreach (string path in temporaryKeys)
                {
                    string keyPath = path;
                    CleanupEscape(delegate
                    {
                        using (var registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                        {
                            registry.DeleteSubKeyTree(keyPath, false);
                            using (var remaining = registry.OpenSubKey(keyPath))
                                if (remaining != null) throw new InvalidOperationException("temporary registry key remains");
                        }
                    }, "remove " + keyPath, cleanupErrors);
                }
                CleanupEscape(delegate { app.COMAddIns.Update(); }, "refresh Office add-in list", cleanupErrors);
            }
            CleanupEscape(delegate
            {
                SelectInstalledMode(AutomationElement.FromHandle(new IntPtr(app.HWND)), installedAutomation, originalMode);
                if (!addin.Connect || Status(installedAutomation).Split('|')[0] != originalMode.ToString())
                    throw new InvalidOperationException("installed add-in mode or connection was not restored");
            }, "restore installed mode", cleanupErrors);
            CleanupEscape(delegate { if (previous != null) previous.Activate(); }, "restore previous window", cleanupErrors);
            CleanupEscape(delegate { Cursor.Position = pointer; }, "restore pointer", cleanupErrors);
            CleanupEscape(delegate { finalCount = app.Presentations.Count; }, "read remaining presentation count", cleanupErrors);
            Console.WriteLine("Temporary document removed; presentations=" + finalCount + ", original=" + originalCount);
            if (ownApplication && finalCount == 0)
                CleanupEscape(delegate { if (app.SlideShowWindows.Count == 0) app.Quit(); }, "close test-created empty PowerPoint", cleanupErrors);
            foreach (string error in cleanupErrors) Console.Error.WriteLine("Cleanup failed: " + error);
        }
        Equal(finalCount, originalCount, "user presentation count preserved");
        Check(cleanupErrors.Count == 0, "all cleanup and restoration operations succeeded: " + string.Join("; ", cleanupErrors));
        Check(failures.Count == 0, string.Join("; ", failures));
    }
}
