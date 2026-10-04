using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace SlidePace
{
    internal sealed class OverlaySurface
    {
        public IntPtr Window;
        public Rectangle Bounds;
        public bool Presenter;
        public string Device;
    }

    internal static class ShowWindows
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct NativeRect { public int Left, Top, Right, Bottom; }
        [StructLayout(LayoutKind.Sequential)]
        private struct NativePoint { public int X, Y; }
        private delegate bool WindowVisitor(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out NativeRect rectangle);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref NativePoint point);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr window, out NativeRect rectangle);
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
        [DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr window, uint command);
        [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
        [DllImport("user32.dll")] private static extern bool EnumWindows(WindowVisitor visitor, IntPtr parameter);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetClassName(IntPtr window, StringBuilder name, int count);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder name, int count);

        internal static bool TryGetBounds(IntPtr window, out Rectangle bounds)
        {
            bounds = Rectangle.Empty;
            NativeRect client;
            if (window == IntPtr.Zero || !IsWindow(window) || !IsWindowVisible(window) || IsIconic(window) || IsIconic(Root(window)) || !GetClientRect(window, out client)) return false;
            var origin = new NativePoint();
            if (!ClientToScreen(window, ref origin)) return false;
            bounds = new Rectangle(origin.X, origin.Y, client.Right - client.Left, client.Bottom - client.Top);
            return bounds.Width >= 160 && bounds.Height >= 100;
        }

        private static bool CanDisplay(IntPtr window, Rectangle bounds)
        {
            IntPtr foreground = GetForegroundWindow();
            if (foreground == IntPtr.Zero) return true;
            IntPtr showRoot = Root(window);
            IntPtr current = Root(foreground); // GA_ROOT: do not treat the editing owner as the show.
            for (int step = 0; current != IntPtr.Zero && step < 12; step++, current = GetWindow(current, 4))
                if (current == window || current == showRoot) return true;
            NativeRect front;
            if (!GetWindowRect(foreground, out front)) return true;
            // A show on the other monitor remains visible. An editor or another app
            // in front of this surface suppresses its overlay without pausing time.
            return !bounds.IntersectsWith(Rectangle.FromLTRB(front.Left, front.Top, front.Right, front.Bottom));
        }

        internal static IntPtr Root(IntPtr window) { return GetAncestor(window, 2); }

        internal static string Describe(IntPtr window)
        {
            var name = new StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            Rectangle bounds;
            return "window=" + window + ", class=" + name + ", visible=" + IsWindowVisible(window) +
                ", minimized=" + IsIconic(window) + ", client=" + (TryGetBounds(window, out bounds) ? bounds.ToString() : "unavailable") +
                ", root=" + GetAncestor(window, 2) + ", foreground=" + GetForegroundWindow() + ", foregroundRoot=" + GetAncestor(GetForegroundWindow(), 2);
        }

        internal static List<OverlaySurface> Capture(IntPtr host, IntPtr show)
        {
            var result = new List<OverlaySurface>();
            Rectangle audience;
            if (!TryGetBounds(show, out audience)) return result;
            if (CanDisplay(show, audience)) Add(result, show, audience, false);
            uint process;
            GetWindowThreadProcessId(host != IntPtr.Zero ? host : show, out process);
            EnumWindows(delegate(IntPtr window, IntPtr parameter)
            {
                if (window == show) return true;
                uint candidate;
                GetWindowThreadProcessId(window, out candidate);
                if (candidate != process) return true;
                var name = new StringBuilder(256);
                GetClassName(window, name, name.Capacity);
                if (name.ToString() == "PPTFrameClass" || name.ToString() == "#32770") return true;
                bool presenter = name.ToString().IndexOf("Presenter", StringComparison.OrdinalIgnoreCase) >= 0;
                if (!presenter)
                {
                    name.Clear();
                    GetWindowText(window, name, name.Capacity);
                    presenter = name.ToString().IndexOf("Presenter View", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.ToString().IndexOf("演示者视图", StringComparison.Ordinal) >= 0;
                }
                Rectangle area;
                if (presenter && TryGetBounds(window, out area) && CanDisplay(window, area)) Add(result, window, area, true);
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private static void Add(List<OverlaySurface> surfaces, IntPtr window, Rectangle area, bool presenter)
        {
            Screen screen = Screen.FromRectangle(area);
            area.Intersect(screen.Bounds);
            if (area.Width < 160 || area.Height < 100) return;
            if (surfaces.Exists(delegate(OverlaySurface surface) { return surface.Bounds == area; })) return;
            surfaces.Add(new OverlaySurface { Window = Root(window), Bounds = area, Presenter = presenter, Device = screen.DeviceName });
        }
    }
}
