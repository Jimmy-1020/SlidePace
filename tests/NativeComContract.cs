using System;
using System.Runtime.InteropServices;
using SlidePace;

[ComVisible(true)]
[Guid("A96AD16C-981F-4DE7-A96E-D78B504CD272")]
[ClassInterface(ClassInterfaceType.None)]
public sealed class ConnectionProbe : IDTExtensibility2
{
    public int Calls;
    public int Mode;
    public Array LastArray;
    public void OnConnection(object host, int mode, object instance, ref Array custom) { Calls++; Mode = mode; LastArray = custom; }
    public void OnDisconnection(int mode, ref Array custom) { Calls++; Mode = mode; LastArray = custom; }
    public void OnAddInsUpdate(ref Array custom) { Calls++; LastArray = custom; }
    public void OnStartupComplete(ref Array custom) { Calls++; LastArray = custom; }
    public void OnBeginShutdown(ref Array custom) { Calls++; LastArray = custom; }
}

internal static partial class Program
{
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int NativeConnect(IntPtr self, IntPtr host, int mode, IntPtr instance, ref IntPtr custom);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int NativeDisconnect(IntPtr self, int mode, ref IntPtr custom);
    [UnmanagedFunctionPointer(CallingConvention.StdCall)]
    private delegate int NativeCustom(IntPtr self, ref IntPtr custom);
    [DllImport("oleaut32.dll")] private static extern IntPtr SafeArrayCreateVector(ushort type, int lowerBound, uint count);
    [DllImport("oleaut32.dll")] private static extern int SafeArrayPutElement(IntPtr array, ref int index, IntPtr variant);
    [DllImport("oleaut32.dll")] private static extern int SafeArrayDestroy(IntPtr array);
    [DllImport("oleaut32.dll")] private static extern int VariantClear(IntPtr variant);
    private static void NativeComContract()
    {
        foreach (var method in typeof(IDTExtensibility2).GetMethods())
        {
            var parameter = method.GetParameters()[method.GetParameters().Length - 1];
            var attribute = (MarshalAsAttribute)Attribute.GetCustomAttribute(parameter, typeof(MarshalAsAttribute));
            Check(parameter.IsIn && attribute != null && attribute.Value == UnmanagedType.SafeArray && attribute.SafeArraySubType == VarEnum.VT_VARIANT,
                "native SAFEARRAY(VARIANT) contract: " + method.Name);
        }
        var probe = new ConnectionProbe();
        IntPtr pointer = Marshal.GetComInterfaceForObject(probe, typeof(IDTExtensibility2));
        IntPtr array = SafeArrayCreateVector((ushort)VarEnum.VT_VARIANT, 0, 2);
        Check(array != IntPtr.Zero, "native array allocated");
        try
        {
            object[] values = { "native-office-caller", 42 };
            for (int index = 0; index < values.Length; index++)
            {
                IntPtr variant = Marshal.AllocCoTaskMem(32);
                try
                {
                    Marshal.GetNativeVariantForObject(values[index], variant);
                    int position = index;
                    Marshal.ThrowExceptionForHR(SafeArrayPutElement(array, ref position, variant));
                    VariantClear(variant);
                }
                finally { Marshal.FreeCoTaskMem(variant); }
            }
            IntPtr table = Marshal.ReadIntPtr(pointer);
            var connect = (NativeConnect)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table, IntPtr.Size * 7), typeof(NativeConnect));
            Marshal.ThrowExceptionForHR(connect(pointer, IntPtr.Zero, 1, IntPtr.Zero, ref array));
            Equal(probe.Mode, 1, "native connect enum received");
            Equal(probe.LastArray.GetValue(0), "native-office-caller", "native variant string marshaled");
            Equal(probe.LastArray.GetValue(1), 42, "native variant number marshaled");
            var disconnect = (NativeDisconnect)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table, IntPtr.Size * 8), typeof(NativeDisconnect));
            Marshal.ThrowExceptionForHR(disconnect(pointer, 0, ref array));
            for (int method = 9; method <= 11; method++)
            {
                var callback = (NativeCustom)Marshal.GetDelegateForFunctionPointer(Marshal.ReadIntPtr(table, IntPtr.Size * method), typeof(NativeCustom));
                Marshal.ThrowExceptionForHR(callback(pointer, ref array));
                Equal(probe.LastArray.Length, 2, "native lifecycle array received");
            }
            Equal(probe.Calls, 5, "all five lifecycle methods invoked through unmanaged vtable");
        }
        finally
        {
            SafeArrayDestroy(array);
            Marshal.Release(pointer);
        }
    }
}
