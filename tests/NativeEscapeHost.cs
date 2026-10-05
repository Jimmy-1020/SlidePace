using System;
using System.IO;
using System.Collections;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using SlidePace;
using Office = Microsoft.Office.Core;

// A separate temporary COM identity lets Office load the exact production DLL
// in-process without replacing an installed SlidePace registration or its data.
[ComVisible(true)]
[Guid("86B4DB30-2296-4B59-938D-11A06F843410")]
[ProgId("SlidePace.EscapeValidation")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class NativeEscapeHost : IDTExtensibility2
{
    private readonly object plugin;
    private System.Windows.Forms.Control dispatcher;
    public NativeEscapeHost()
    {
        string directory = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
            "../../../artifacts/esc-investigation/native-config-" + Guid.NewGuid().ToString("N")));
        // Weak-named .NET Framework assemblies may bind to Office's already
        // loaded installed version. LoadFile isolates the exact newly built DLL.
        var library = Assembly.LoadFile(Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "SlidePace.AddIn.dll"));
        plugin = Activator.CreateInstance(library.GetType("SlidePace.PowerPointAddIn"),
            BindingFlags.Instance | BindingFlags.NonPublic, null, new object[] { directory }, null);
    }
    private object Invoke(string name, params object[] args)
    {
        return OnHostThread(delegate
        {
            return plugin.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                null, plugin, args);
        });
    }
    private object OnHostThread(Func<object> action)
    {
        // A managed COM automation call can arrive on an RPC worker thread.
        // Apply settings on the Office UI thread, as the real settings dialog does.
        if (dispatcher != null && dispatcher.InvokeRequired) return dispatcher.Invoke(action);
        return action();
    }
    public void OnConnection(object host, int connectMode, object instance, ref Array custom)
    {
        dispatcher = new System.Windows.Forms.Control();
        IntPtr handle = dispatcher.Handle;
        object[] args = { host, connectMode, null, custom };
        Invoke("OnConnection", args);
        custom = (Array)args[3];
        var addin = instance as Office.COMAddIn;
        if (addin != null) addin.Object = this;
    }
    public void OnDisconnection(int removeMode, ref Array custom)
    {
        object[] args = { removeMode, custom };
        Invoke("OnDisconnection", args);
        custom = (Array)args[1];
        OnHostThread(delegate { if (dispatcher != null) dispatcher.Dispose(); return null; });
        dispatcher = null;
    }
    public void OnAddInsUpdate(ref Array custom) { object[] args = { custom }; Invoke("OnAddInsUpdate", args); custom = (Array)args[0]; }
    public void OnStartupComplete(ref Array custom) { object[] args = { custom }; Invoke("OnStartupComplete", args); custom = (Array)args[0]; }
    public void OnBeginShutdown(ref Array custom) { object[] args = { custom }; Invoke("OnBeginShutdown", args); custom = (Array)args[0]; }
    public string GetDiagnosticStatus() { return (string)Invoke("GetDiagnosticStatus"); }
    public string GetDiagnosticAssembly() { return plugin.GetType().Assembly.FullName + "|" + plugin.GetType().Assembly.Location; }
    public void SetModeForValidation(int mode)
    {
        var modeType = plugin.GetType().Assembly.GetType("SlidePace.TimerMode");
        Invoke("SelectMode", Enum.ToObject(modeType, mode));
    }
    public void SetPresenterTimerForValidation(bool visible)
    {
        OnHostThread(delegate
        {
            var settings = plugin.GetType().GetProperty("Settings", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(plugin, null);
            settings.GetType().GetField("ShowPresenterTimer").SetValue(settings, visible);
            Invoke("ApplySettings", settings);
            return null;
        });
    }
    public double GetElapsedForValidation()
    {
        return (double)OnHostThread(delegate
        {
            var engine = plugin.GetType().GetProperty("Engine", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(plugin, null);
            var snapshot = engine.GetType().GetMethod("Snapshot").Invoke(engine, null);
            return snapshot.GetType().GetProperty("ElapsedSeconds").GetValue(snapshot, null);
        });
    }
    public string GetOverlayDetailsForValidation()
    {
        return (string)OnHostThread(delegate
        {
            var manager = plugin.GetType().GetProperty("Overlays", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(plugin, null);
            var forms = (IEnumerable)manager.GetType().GetProperty("Forms").GetValue(manager, null);
            return string.Join("; ", forms.Cast<System.Windows.Forms.Form>().Select(form =>
                "handle=" + (form.IsHandleCreated ? form.Handle.ToString() : "none") +
                " visible=" + form.Visible + " disposed=" + form.IsDisposed + " text=" + form.Text + " bounds=" + form.Bounds));
        });
    }
}
