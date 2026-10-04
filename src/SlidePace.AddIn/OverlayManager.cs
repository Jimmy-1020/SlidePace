using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SlidePace
{
    internal sealed class OverlayManager : IDisposable
    {
        private readonly List<OverlayForm> forms = new List<OverlayForm>();
        private UserSettings settings;
        private readonly Action settingsChanged;
        private readonly Action startPause;
        private readonly Action resume;
        private readonly Action reset;
        private readonly Action openSettings;
        private readonly Action<TimerMode> modeChanged;
        private string signature = "";
        private bool collapsed;
        private IntPtr hostWindow;
        private IntPtr slideShowWindow;
        public IList<OverlayForm> Forms { get { return forms.AsReadOnly(); } }

        public OverlayManager(UserSettings initialSettings, Action persist, Action onStartPause, Action onContinue,
            Action onReset, Action onSettings, Action<TimerMode> onModeChanged)
        {
            settings = initialSettings;
            settingsChanged = persist;
            startPause = onStartPause;
            resume = onContinue;
            reset = onReset;
            openSettings = onSettings;
            modeChanged = onModeChanged;
        }

        public void SetHostWindows(IntPtr host, IntPtr show)
        {
            if (hostWindow == host && slideShowWindow == show) return;
            hostWindow = host; slideShowWindow = show; signature = "";
        }
        public void ApplySettings(UserSettings value) { settings = value; signature = ""; }
        public void ResetCollapsed() { collapsed = false; signature = ""; }

        private Screen Find(string device, Screen fallback)
        {
            return Screen.AllScreens.FirstOrDefault(delegate(Screen screen) { return screen.DeviceName == device; }) ?? fallback;
        }

        public void Render(TimerSnapshot snapshot, bool active)
        {
            if (!active || snapshot.Mode == TimerMode.None)
            {
                CloseForms();
                return;
            }
            List<OverlaySurface> surfaces;
            if (slideShowWindow != IntPtr.Zero || hostWindow != IntPtr.Zero) surfaces = ShowWindows.Capture(hostWindow, slideShowWindow);
            else
            {
                // Standalone UI validation; production always supplies the actual show HWND.
                Screen audience = Find(settings.AudienceDevice, Screen.PrimaryScreen);
                Screen presenter = Find(settings.PresenterDevice, Screen.PrimaryScreen);
                if (presenter.DeviceName == audience.DeviceName && Screen.AllScreens.Length > 1)
                    presenter = Screen.AllScreens.First(delegate(Screen screen) { return screen.DeviceName != audience.DeviceName; });
                surfaces = new List<OverlaySurface> { new OverlaySurface { Bounds = presenter.Bounds, Device = presenter.DeviceName, Presenter = true } };
                if (presenter.DeviceName != audience.DeviceName) surfaces.Add(new OverlaySurface { Bounds = audience.Bounds, Device = audience.DeviceName });
            }
            if (surfaces.Count == 0) { CloseForms(); return; }
            string desired = string.Join(";", surfaces.Select(delegate(OverlaySurface surface) { return surface.Window + ":" + surface.Bounds + ":" + surface.Presenter; })) + "|" + settings.FontSize + "|" + collapsed + "|" +
                settings.PresenterPosition + "|" + settings.AudiencePosition + "|" +
                string.Join(";", Screen.AllScreens.Select(delegate(Screen screen) { return screen.DeviceName + screen.Bounds; }));
            if (desired != signature)
            {
                CloseForms();
                foreach (OverlaySurface surface in surfaces) CreateForm(surface, snapshot);
                signature = desired;
            }
            // One snapshot is shared by all windows, including its wall-clock value.
            foreach (OverlayForm form in forms)
            {
                form.Render(snapshot);
                form.KeepAboveShow();
            }
        }

        private void CreateForm(OverlaySurface surface, TimerSnapshot snapshot)
        {
            var form = new OverlayForm(surface.Device, surface.Presenter);
            form.DisplayArea = surface.Bounds;
            // Full-screen PowerPoint surfaces can themselves be topmost. Keep the
            // overlay above them; Capture restricts its lifetime to visible show views.
            form.TopMost = true;
            form.ApplySettings(settings, collapsed);
            form.StartPauseRequested += startPause;
            form.ContinueRequested += resume;
            form.ResetRequested += reset;
            form.SettingsRequested += openSettings;
            form.ModeRequested += modeChanged;
            form.CollapseRequested += delegate { collapsed = !collapsed; signature = ""; };
            form.DragFinished += SaveLocation;
            form.Render(snapshot);
            form.Location = GetLocation(surface.Bounds, surface.Device, form.Size, surface.Presenter);
            forms.Add(form);
            if (surface.Window != IntPtr.Zero) form.Show(new NativeOwner(surface.Window));
            else form.Show();
        }

        private Point GetLocation(Rectangle area, string device, Size size, bool presenter)
        {
            OverlayPosition position = presenter ? settings.PresenterPosition : settings.AudiencePosition;
            int margin = Math.Max(12, (int)(size.Width / 360.0 * 18));
            if (position == OverlayPosition.Custom)
            {
                SavedLocation custom = settings.Locations.FirstOrDefault(delegate(SavedLocation item) { return item.Device == device; });
                if (custom != null)
                    return OverlayForm.Clamp(new Point(area.Left + (int)Math.Round(custom.X * Math.Max(0, area.Width - size.Width)),
                        area.Top + (int)Math.Round(custom.Y * Math.Max(0, area.Height - size.Height))), area, size);
            }
            bool left = position == OverlayPosition.TopLeft || position == OverlayPosition.BottomLeft;
            bool bottom = position == OverlayPosition.BottomLeft || position == OverlayPosition.BottomRight;
            return OverlayForm.Clamp(new Point(left ? area.Left + margin : area.Right - size.Width - margin,
                bottom ? area.Bottom - size.Height - margin : area.Top + margin), area, size);
        }

        private void SaveLocation(OverlayForm form)
        {
            Rectangle area = form.AreaBounds;
            Rectangle compact = form.CompactBounds;
            settings.Locations.RemoveAll(delegate(SavedLocation location) { return location.Device == form.DeviceName; });
            settings.Locations.Add(new SavedLocation
            {
                Device = form.DeviceName,
                X = (double)(compact.Left - area.Left) / Math.Max(1, area.Width - compact.Width),
                Y = (double)(compact.Top - area.Top) / Math.Max(1, area.Height - compact.Height)
            });
            if (form.IsPresenter) settings.PresenterPosition = OverlayPosition.Custom;
            else settings.AudiencePosition = OverlayPosition.Custom;
            // Preserve the visible form while dragging; the next settings change rebuilds it.
            signature = "";
            settingsChanged();
        }

        public void CloseForms()
        {
            foreach (OverlayForm form in forms.ToArray()) { form.Close(); form.Dispose(); }
            forms.Clear();
            signature = "";
        }
        public void Dispose() { CloseForms(); }
    }
}
