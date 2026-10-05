using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SlidePace
{
    [ComVisible(true)]
    [Guid("B18A80F9-540D-4F9A-9F1D-A4798AC2A398")]
    [ProgId("SlidePace.PowerPointAddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class PowerPointAddIn : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private readonly SettingsStore store;
        private PowerPoint.Application application;
        private PowerPoint.Presentation activePresentation;
        private PowerPoint.SlideShowWindow activeShow;
        private Office.IRibbonUI ribbon;
        private TimerEngine engine;
        private UserSettings settings;
        private OverlayManager overlays;
        private System.Windows.Forms.Timer refresh;
        private bool disposed;
        private bool eventsConnected;
        private int pollTicks;
        private string previousState;
        private string notice;
        internal TimerEngine Engine { get { return engine; } }
        internal OverlayManager Overlays { get { return overlays; } }
        internal UserSettings Settings { get { return settings; } }
        internal int BeginEventCount { get; private set; }
        internal int EndEventCount { get; private set; }

        public PowerPointAddIn() : this(DefaultDataDirectory()) { }
        internal PowerPointAddIn(string directory) { store = new SettingsStore(directory); }

        private static string DefaultDataDirectory()
        {
#if DEBUG
            return Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "validation-data");
#else
            return SettingsStore.DefaultDirectory;
#endif
        }

        public void OnConnection(object host, int connectMode, object instance, ref Array custom)
        {
            Guard(delegate
            {
                if (application != null) return;
                disposed = false;
                application = (PowerPoint.Application)host;
                settings = store.Load();
                engine = new TimerEngine(new SystemTimeSource());
                engine.SetCountdownSeconds(settings.CountdownSeconds);
                engine.SetCountdownEndBehavior(settings.CountdownEndBehavior);
                engine.SelectMode((TimerMode)settings.Mode);
                overlays = new OverlayManager(settings, Persist, StartPause, Continue, Reset, OpenSettings,
                    delegate(TimerMode mode) { SelectMode(engine.Mode == mode ? TimerMode.None : mode); });
                application.SlideShowBegin += SlideShowBegin;
                application.SlideShowEnd += SlideShowEnd;
                application.PresentationClose += PresentationClose;
                eventsConnected = true;
                refresh = new System.Windows.Forms.Timer { Interval = 100 };
                refresh.Tick += RefreshTick;
                refresh.Start();
                var addin = instance as Office.COMAddIn;
                if (addin != null) addin.Object = this;
                PollSlideShow();
                InvalidateRibbon();
            }, false);
        }

        public void OnDisconnection(int removeMode, ref Array custom) { Disconnect(); }
        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { Disconnect(); }

        public string GetDiagnosticStatus()
        {
            if (engine == null) return "not-connected";
            TimerSnapshot value = engine.Snapshot();
            return engine.Mode + "|" + value.Text + "|" + value.IsRunning + "|" +
                (overlays == null ? 0 : overlays.Forms.Count) + "|" + BeginEventCount + "|" + EndEventCount + "|" + (ribbon != null);
        }

        private void Disconnect()
        {
            if (disposed) return;
            disposed = true;
            if (refresh != null) { refresh.Stop(); refresh.Dispose(); refresh = null; }
            if (engine != null) engine.EndSlideShow();
            if (overlays != null) { overlays.Dispose(); overlays = null; }
            if (eventsConnected && application != null)
            {
                try { application.SlideShowBegin -= SlideShowBegin; } catch { }
                try { application.SlideShowEnd -= SlideShowEnd; } catch { }
                try { application.PresentationClose -= PresentationClose; } catch { }
            }
            eventsConnected = false;
            activeShow = null;
            activePresentation = null;
            application = null;
            ribbon = null;
            // Office owns these RCWs. Releasing shared application RCWs here can break other add-ins.
        }

        public string GetCustomUI(string ribbonId)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("SlidePace.Ribbon.xml"))
            using (var reader = new StreamReader(stream)) return reader.ReadToEnd();
        }

        public void OnRibbonLoad(Office.IRibbonUI value) { ribbon = value; InvalidateRibbon(); }
        public bool GetModePressed(Office.IRibbonControl control) { return engine != null && engine.Mode == ControlMode(control.Id); }
        public string GetStartLabel(Office.IRibbonControl control) { return engine != null && engine.Snapshot().IsRunning ? "暂停" : "开始"; }
        public bool GetStartEnabled(Office.IRibbonControl control)
        {
            if (!CanControl()) return false;
            TimerSnapshot value = engine.Snapshot();
            return value.IsRunning || !value.HasStarted;
        }
        public bool GetContinueEnabled(Office.IRibbonControl control)
        {
            if (!CanControl()) return false;
            TimerSnapshot value = engine.Snapshot();
            return value.HasStarted && !value.IsRunning && !value.IsCompleted;
        }
        public bool GetResetEnabled(Office.IRibbonControl control)
        {
            return engine != null && (engine.Mode == TimerMode.CountUp || engine.Mode == TimerMode.CountDown);
        }
        private bool CanControl()
        {
            return engine != null && engine.IsSlideShowActive && (engine.Mode == TimerMode.CountUp || engine.Mode == TimerMode.CountDown);
        }
        private static TimerMode ControlMode(string id)
        {
            return id == "ModeCountUp" ? TimerMode.CountUp : id == "ModeCountDown" ? TimerMode.CountDown : id == "ModeClock" ? TimerMode.Clock : TimerMode.None;
        }
        public void OnModeChanged(Office.IRibbonControl control, bool pressed) { SelectMode(pressed ? ControlMode(control.Id) : TimerMode.None); }
        public void OnStartPause(Office.IRibbonControl control) { StartPause(); }
        public void OnContinue(Office.IRibbonControl control) { Continue(); }
        public void OnReset(Office.IRibbonControl control) { Reset(); }
        public void OnOpenSettings(Office.IRibbonControl control) { OpenSettings(); }
        public void OnHelp(Office.IRibbonControl control)
        {
            MessageBox.Show("每次打开 PowerPoint 默认不选计时模式。点击一种模式选中，再次点击取消；进入放映自动显示并开始。\n\n框体半透明，平时只显示时间数字；悬停展开按钮，移开恢复紧凑高度。运行时“开始”原位变为“暂停”。\n\n在“计时器设置”中选择时长、归零后停止或继续顺计时、字体、字号、颜色和位置。时间不显示负数；倒计时归零后默认变为红色，可自选颜色。\n\n演讲者和观众屏幕默认同步显示。取消勾选“演示者侧显示计时器”并保存，仅隐藏该侧计时框，保留演示者视图与观众计时。右键可收起框体或切换模式。\n\n取消选中模式后所有屏幕都不显示计时。结束放映自动暂停；再次放映继续本次会话读数。已完成的倒计时保持零，开始新一轮请先重置。",
                "SlidePace 1.0.5 · 使用说明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        internal void SelectMode(TimerMode mode)
        {
            Guard(delegate
            {
                engine.SelectMode(mode);
                settings.Mode = (int)mode;
                Persist();
                if (mode == TimerMode.None) overlays.ResetCollapsed();
                Render();
                InvalidateRibbon();
            }, true);
        }
        internal void StartPause()
        {
            Guard(delegate
            {
                if (!CanControl()) return;
                if (engine.Snapshot().IsRunning) engine.Pause(); else if (!engine.Snapshot().HasStarted) engine.Start();
                Render();
                InvalidateRibbon();
            }, true);
        }
        internal void Continue()
        {
            Guard(delegate { if (CanControl()) engine.Start(); Render(); InvalidateRibbon(); }, true);
        }
        internal void Reset()
        {
            Guard(delegate { engine.Reset(); Render(); InvalidateRibbon(); }, true);
        }
        internal void ApplySettings(UserSettings value)
        {
            value.Validate();
            engine.SetCountdownSeconds(value.CountdownSeconds);
            engine.SetCountdownEndBehavior(value.CountdownEndBehavior);
            engine.SelectMode((TimerMode)value.Mode);
            settings = value;
            overlays.ApplySettings(settings);
            Persist();
            Render();
            InvalidateRibbon();
        }
        private void OpenSettings()
        {
            Guard(delegate
            {
                bool countdownRunning = engine.Mode == TimerMode.CountDown && engine.Snapshot().IsRunning;
                using (var dialog = new SettingsForm(settings, countdownRunning))
                {
                    DialogResult result = application == null ? dialog.ShowDialog() : dialog.ShowDialog(new NativeOwner(new IntPtr(application.HWND)));
                    if (result == DialogResult.OK) ApplySettings(dialog.Result);
                }
            }, true);
        }
        private void Persist()
        {
            try { store.Save(settings); }
            catch (Exception error) { store.Log(error); notice = "设置未保存：" + error.Message; }
        }
        private void InvalidateRibbon()
        {
            if (ribbon != null) try { ribbon.Invalidate(); } catch (COMException) { }
        }
        private void RefreshTick(object sender, EventArgs args)
        {
            Guard(delegate
            {
                if (++pollTicks % 5 == 0) PollSlideShow();
                Render();
            }, false);
        }
        private void Render()
        {
            if (engine == null || overlays == null) return;
            TimerSnapshot value = engine.Snapshot();
            if (!string.IsNullOrEmpty(notice) && !value.IsOvertime) value.Status = notice;
            overlays.Render(value, engine.IsSlideShowActive);
            string state = value.Mode + ":" + value.IsRunning + ":" + value.HasStarted + ":" + value.IsCompleted;
            if (previousState != state) { previousState = state; InvalidateRibbon(); }
        }
        private void SlideShowBegin(PowerPoint.SlideShowWindow window)
        {
            BeginEventCount++;
            Guard(delegate { BeginShow(window); }, false);
        }
        private void BeginShow(PowerPoint.SlideShowWindow window)
        {
            if (engine.IsSlideShowActive)
            {
                if (!SameComObject(activePresentation, window.Presentation)) notice = "计时仍属于第一场放映";
                return;
            }
            activeShow = window;
            activePresentation = window.Presentation;
            notice = null;
            overlays.SetHostWindows(new IntPtr(application.HWND), new IntPtr(window.HWND));
            overlays.ResetCollapsed();
            engine.BeginSlideShow();
            // Let PowerPoint finish constructing its full-screen and presenter
            // windows before the next refresh creates owned overlay windows.
            InvalidateRibbon();
        }
        private void SlideShowEnd(PowerPoint.Presentation presentation)
        {
            EndEventCount++;
            Guard(delegate { if (SameComObject(activePresentation, presentation)) EndShow(); }, false);
        }
        private void PresentationClose(PowerPoint.Presentation presentation)
        {
            Guard(delegate { if (SameComObject(activePresentation, presentation)) EndShow(); }, false);
        }
        private void EndShow()
        {
            engine.EndSlideShow();
            activeShow = null;
            activePresentation = null;
            notice = null;
            overlays.CloseForms();
            InvalidateRibbon();
        }
        private void PollSlideShow()
        {
            if (application == null || disposed) return;
            try
            {
                PowerPoint.SlideShowWindows windows = application.SlideShowWindows;
                if (engine.IsSlideShowActive)
                {
                    bool found = false;
                    for (int index = 1; index <= windows.Count; index++)
                    {
                        PowerPoint.SlideShowWindow candidate = windows[index];
                        if (!SameComObject(candidate.Presentation, activePresentation)) continue;
                        if (candidate.View.State == PowerPoint.PpSlideShowState.ppSlideShowDone) continue;
                        activeShow = candidate;
                        overlays.SetHostWindows(new IntPtr(application.HWND), new IntPtr(candidate.HWND));
                        found = true;
                        break;
                    }
                    if (!found) EndShow();
                }
                else
                {
                    for (int index = 1; index <= windows.Count; index++)
                    {
                        PowerPoint.SlideShowWindow candidate = windows[index];
                        System.Drawing.Rectangle area;
                        if (candidate.View.State != PowerPoint.PpSlideShowState.ppSlideShowDone && ShowWindows.TryGetBounds(new IntPtr(candidate.HWND), out area))
                        { BeginShow(candidate); break; }
                    }
                }
            }
            catch (COMException error) { store.Log(error); }
        }
        private static bool SameComObject(object left, object right)
        {
            if (left == null || right == null) return false;
            if (ReferenceEquals(left, right)) return true;
            IntPtr first = IntPtr.Zero;
            IntPtr second = IntPtr.Zero;
            try
            {
                first = Marshal.GetIUnknownForObject(left);
                second = Marshal.GetIUnknownForObject(right);
                return first == second;
            }
            finally
            {
                if (first != IntPtr.Zero) Marshal.Release(first);
                if (second != IntPtr.Zero) Marshal.Release(second);
            }
        }
        private void Guard(Action action, bool userInitiated)
        {
            try { action(); }
            catch (Exception error)
            {
                store.Log(error);
                if (userInitiated) MessageBox.Show(error.Message, "SlidePace", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else notice = "计时暂不可用，请检查插件设置";
            }
        }
    }
}
