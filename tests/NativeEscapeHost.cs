using System;
using System.IO;
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
        return plugin.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
            null, plugin, args);
    }
    public void OnConnection(object host, int connectMode, object instance, ref Array custom)
    {
        object[] args = { host, connectMode, null, custom };
        Invoke("OnConnection", args);
        custom = (Array)args[3];
        var addin = instance as Office.COMAddIn;
        if (addin != null) addin.Object = this;
    }
    public void OnDisconnection(int removeMode, ref Array custom) { object[] args = { removeMode, custom }; Invoke("OnDisconnection", args); custom = (Array)args[1]; }
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
}
