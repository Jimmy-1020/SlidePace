using System;
using System.Reflection;
using System.Runtime.InteropServices;

internal static class InspectInterop
{
    private static void Main(string[] args)
    {
        foreach (string path in args)
        {
            Assembly assembly = Assembly.LoadFrom(path);
            foreach (Type type in assembly.GetTypes())
            {
                if (type.Name != "IDTExtensibility2" && type.Name != "IRibbonExtensibility" &&
                    type.Name != "IRibbonControl" && type.Name != "IRibbonUI") continue;
                Console.WriteLine(type.FullName + " " + type.GUID);
                foreach (object attribute in type.GetCustomAttributes(false))
                {
                    InterfaceTypeAttribute kind = attribute as InterfaceTypeAttribute;
                    if (kind != null) Console.WriteLine("  " + kind.Value);
                }
                foreach (MethodInfo method in type.GetMethods())
                {
                    Console.WriteLine("  " + method);
                    foreach (ParameterInfo parameter in method.GetParameters())
                    {
                        Console.WriteLine("    " + parameter.Name + " " + parameter.Attributes);
                        var marshal = (MarshalAsAttribute)Attribute.GetCustomAttribute(parameter, typeof(MarshalAsAttribute));
                        if (marshal != null) Console.WriteLine("      MarshalAs=" + marshal.Value + " SafeArraySubType=" + marshal.SafeArraySubType);
                    }
                }
            }
        }
    }
}
