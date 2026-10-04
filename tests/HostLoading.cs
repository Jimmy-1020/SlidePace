using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Automation;
using Microsoft.Win32;
using SlidePace.Setup;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

internal static partial class Program
{
    private static string Status(object automation)
    {
        return (string)automation.GetType().InvokeMember("GetDiagnosticStatus", BindingFlags.InvokeMethod, null, automation, new object[0]);
    }
    private static void HostLoading()
    {
        string debugLibrary = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "../../../src/SlidePace.AddIn/bin/Debug/SlidePace.AddIn.dll"));
        Check(File.Exists(debugLibrary), "debug plugin exists; its data stays inside the workspace");
        var options = new InstallOptions
        {
            Root = Path.Combine(root, "temporary-registration-" + Guid.NewGuid().ToString("N")),
            TestMode = true,
            PayloadOverrideForTests = File.ReadAllBytes(debugLibrary)
        };
        var installer = new InstallerEngine(options, null);
        string[] keys = { installer.ClassKey, installer.ProgIdKey, installer.AddInKey, installer.UninstallKey };
        using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
            foreach (string key in keys) using (RegistryKey existing = registry.OpenSubKey(key))
                Check(existing == null, "no existing SlidePace installation will be overwritten");
        int[] existingProcesses = Process.GetProcessesByName("POWERPNT").Select(delegate(Process process) { return process.Id; }).ToArray();
        PowerPoint.Application app;
        try { app = (PowerPoint.Application)Marshal.GetActiveObject("PowerPoint.Application"); }
        catch (COMException) { app = (PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application")); }
        uint hostProcess;
        GetWindowThreadProcessId(new IntPtr(app.HWND), out hostProcess);
        bool ownHost = !existingProcesses.Contains((int)hostProcess);
        if (app.Presentations.Count != 0) throw new InvalidOperationException("Host-load verification requires an empty PowerPoint instance so user documents are never exposed to loading failures.");
        if (app.SlideShowWindows.Count > 0) throw new InvalidOperationException("Existing slide show detected; test aborted.");
        PowerPoint.DocumentWindow previousWindow = null;
        try { previousWindow = app.ActiveWindow; } catch { }
        int count = app.Presentations.Count;
        Office.COMAddIn addin = null;
        PowerPoint.Presentation presentation = null;
        AutomationElement previousTab = null;
        try
        {
            app.Visible = Office.MsoTriState.msoTrue;
            presentation = app.Presentations.Add(Office.MsoTriState.msoTrue);
            presentation.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutBlank);
            installer.Install();
            app.COMAddIns.Update();
            object name = InstallerEngine.ProgId;
            addin = app.COMAddIns.Item(ref name);
            addin.Connect = true;
            Until(delegate { return addin.Connect && addin.Object != null; }, 10000, "Office itself activated the registered COM add-in");
            object automation = addin.Object;
            Console.WriteLine("Registered host status: " + Status(automation));
            Check(Status(automation) != "not-connected", "actual COM connection initialized");
            Until(delegate { return Status(automation).EndsWith("|True"); }, 5000, "Office requested Ribbon XML and invoked onLoad");
            AutomationElement host = AutomationElement.FromHandle(new IntPtr(app.HWND));
            AutomationElementCollection tabs = host.FindAll(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
            foreach (AutomationElement tab in tabs)
            {
                object pattern;
                if (tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern) && ((SelectionItemPattern)pattern).Current.IsSelected)
                { previousTab = tab; break; }
            }
            AutomationElement pluginTab = host.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "SlidePace 计时"));
            Check(pluginTab != null, "real PowerPoint ribbon contains SlidePace tab");
            SelectAutomationElement(pluginTab);
            Pump(250);
            AutomationElement clearButton = host.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "全部取消 · 不显示"));
            Check(clearButton == null, "actual ribbon omits redundant clear-mode control");
            Check(Status(automation).StartsWith("None|"), "new PowerPoint session starts with no selected mode");
            AutomationElement upButton = host.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "顺计时"));
            Check(upButton != null, "actual ribbon contains count-up mode button");
            ToggleAutomationElement(upButton);
            Pump(250);
            Check(Status(automation).StartsWith("CountUp|"), "actual ribbon callback selects mode");
            ToggleAutomationElement(upButton);
            Pump(150);
            Check(Status(automation).StartsWith("None|"), "clicking selected mode again deselects it");
            ToggleAutomationElement(upButton);
            Pump(150);
            var slide = presentation.Slides[1];
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            slide.Background.Fill.ForeColor.RGB = ColorTranslator.ToOle(Color.FromArgb(18, 32, 51));
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 60, 160, 700, 100);
            title.TextFrame.TextRange.Text = "SlidePace · 插件自主加载与全屏验证";
            title.TextFrame.TextRange.Font.Size = 32;
            title.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);
            presentation.SlideShowSettings.ShowType = PowerPoint.PpSlideShowType.ppShowTypeSpeaker;
            PowerPoint.SlideShowWindow show = presentation.SlideShowSettings.Run();
            Until(delegate { return Status(automation).Split('|')[2] == "True"; }, 5000, "registered plugin auto starts on actual full-screen show");
            Pump(1200);
            string[] running = Status(automation).Split('|');
            Equal(running[3], Math.Min(2, System.Windows.Forms.Screen.AllScreens.Length).ToString(), "registered plugin displays both screen overlays");
            Check(int.Parse(running[4]) > 0, "registered plugin receives actual SlideShowBegin");
            Check(running[1] != "00:00:00", "registered plugin elapsed time advances");
            CheckVisibleHostOverlays(hostProcess);
            show.View.Exit();
            Until(delegate { return Status(automation).Split('|')[2] == "False"; }, 5000, "registered plugin pauses on show exit");
            Equal(Status(automation).Split('|')[3], "0", "registered plugin closes floating windows");
            Check(int.Parse(Status(automation).Split('|')[5]) > 0, "registered plugin receives actual SlideShowEnd");
            Console.WriteLine("Registered full-screen result: " + Status(automation));
            presentation.Saved = Office.MsoTriState.msoTrue;
            presentation.Close();
            presentation = null;
            Equal(app.Presentations.Count, count, "user presentations preserved after host-load test");
        }
        finally
        {
            if (presentation != null)
            {
                try { foreach (PowerPoint.SlideShowWindow show in app.SlideShowWindows) if (show.Presentation == presentation) show.View.Exit(); } catch { }
                try { presentation.Saved = Office.MsoTriState.msoTrue; presentation.Close(); } catch { }
            }
            if (addin != null) try { addin.Connect = false; } catch { }
            using (RegistryKey registry = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64))
                foreach (string key in keys) registry.DeleteSubKeyTree(key, false);
            try { app.COMAddIns.Update(); } catch { }
            try { if (previousWindow != null) previousWindow.Activate(); } catch { }
            try { if (previousTab != null) SelectAutomationElement(previousTab); } catch { }
            if (ownHost) try { if (app.Presentations.Count == 0) app.Quit(); } catch { }
            Console.WriteLine("Temporary Office registration removed. Loaded test DLL copies remain under artifacts until PowerPoint exits.");
        }
    }
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [StructLayout(LayoutKind.Sequential)] private struct WindowRectangle { public int Left, Top, Right, Bottom; }
    private delegate bool WindowVisitor(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(WindowVisitor visitor, IntPtr parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int count);
    [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out WindowRectangle rectangle);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool GetLayeredWindowAttributes(IntPtr window, out uint colorKey, out byte alpha, out uint flags);
    private static void CheckVisibleHostOverlays(uint hostProcess)
    {
        var windows = new List<IntPtr>();
        EnumWindows(delegate(IntPtr window, IntPtr parameter)
        {
            uint process;
            GetWindowThreadProcessId(window, out process);
            if (process == hostProcess && IsWindowVisible(window))
            {
                var text = new StringBuilder(256);
                GetWindowText(window, text, text.Capacity);
                if (text.ToString() == "SlidePace 计时") windows.Add(window);
            }
            return true;
        }, IntPtr.Zero);
        Equal(windows.Count, Math.Min(2, System.Windows.Forms.Screen.AllScreens.Length), "actual full-screen overlay HWNDs are visible");
        int number = 0;
        foreach (IntPtr window in windows)
        {
            WindowRectangle rectangle;
            Check(GetWindowRect(window, out rectangle), "overlay window bounds available");
            uint colorKey, flags;
            byte alpha;
            Check(GetLayeredWindowAttributes(window, out colorKey, out alpha, out flags) && (flags & 2) != 0 && alpha > 0 && alpha < 255,
                "actual overlay uses constant-alpha translucency");
            var location = new Point(rectangle.Left, rectangle.Top);
            var size = new Size(rectangle.Right - rectangle.Left, rectangle.Bottom - rectangle.Top);
            using (var image = new Bitmap(size.Width, size.Height))
            using (Graphics graphics = Graphics.FromImage(image))
            {
                graphics.CopyFromScreen(location, Point.Empty, size);
                Color pixel = image.GetPixel(8, size.Height / 2);
                Color surface = Color.FromArgb(22, 29, 39);
                double opacity = alpha / 255.0;
                Check(new[] { pixel.R, pixel.G, pixel.B }.Zip(new[] { surface.R, surface.G, surface.B },
                    delegate(byte actual, byte foreground) { return actual >= foreground * opacity - 2 && actual <= foreground * opacity + 255 * (1 - opacity) + 2; }).All(delegate(bool valid) { return valid; })
                    && pixel.ToArgb() != Color.FromArgb(18, 32, 51).ToArgb(), "actual full-screen pixels contain the translucent overlay above slideshow");
                image.Save(Path.Combine(root, "fullscreen-overlay-" + (++number) + ".png"), ImageFormat.Png);
            }
        }
    }
    private static void SelectAutomationElement(AutomationElement element)
    {
        object pattern;
        if (element.TryGetCurrentPattern(SelectionItemPattern.Pattern, out pattern)) ((SelectionItemPattern)pattern).Select();
        else if (element.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) ((InvokePattern)pattern).Invoke();
        else throw new InvalidOperationException("Ribbon tab has no selectable automation pattern.");
    }
    private static void ToggleAutomationElement(AutomationElement element)
    {
        object pattern;
        if (element.TryGetCurrentPattern(TogglePattern.Pattern, out pattern)) ((TogglePattern)pattern).Toggle();
        else if (element.TryGetCurrentPattern(InvokePattern.Pattern, out pattern)) ((InvokePattern)pattern).Invoke();
        else throw new InvalidOperationException("Ribbon mode button has no usable automation pattern.");
    }
}
