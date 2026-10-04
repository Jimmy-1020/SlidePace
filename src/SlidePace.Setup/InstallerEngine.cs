using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace SlidePace.Setup
{
    internal sealed class InstallOptions
    {
        public string Root = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "SlidePace");
        public string SettingsRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlidePace");
        public string RegistryPrefix = "";
        public bool TestMode;
        public bool FailAfterRegistryForTest;
        public byte[] PayloadOverrideForTests;
        public string SetupPath = Assembly.GetExecutingAssembly().Location;
    }

    internal sealed class OfficeEnvironment
    {
        public string Executable;
        public string Version;
        public static OfficeEnvironment Detect()
        {
            string executable = null;
            foreach (RegistryView view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
            {
                using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view))
                using (RegistryKey key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE"))
                    if (key != null && File.Exists(Convert.ToString(key.GetValue("")))) executable = Convert.ToString(key.GetValue(""));
                if (executable != null) break;
            }
            if (executable == null) throw new InvalidOperationException("未找到 Microsoft PowerPoint 桌面版。请先安装 PowerPoint，再运行此安装程序。");
            using (var stream = File.OpenRead(executable))
            using (var reader = new BinaryReader(stream))
            {
                stream.Position = 0x3C;
                int offset = reader.ReadInt32();
                stream.Position = offset;
                if (reader.ReadUInt32() != 0x00004550 || reader.ReadUInt16() != 0x8664)
                    throw new InvalidOperationException("当前版本先支持 64 位 PowerPoint。检测到的 PowerPoint 架构不在此安装包的支持范围内。");
            }
            using (RegistryKey root = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
            using (RegistryKey runtime = root.OpenSubKey(@"SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full"))
                if (runtime == null || Convert.ToInt32(runtime.GetValue("Release", 0)) < 528040)
                    throw new InvalidOperationException("需要 .NET Framework 4.8。请安装该 Windows 运行组件后重试；无需安装 Visual Studio。");
            return new OfficeEnvironment { Executable = executable, Version = FileVersionInfo.GetVersionInfo(executable).FileVersion };
        }
    }

    internal sealed class RegistryImage
    {
        internal sealed class Value
        {
            public string Name;
            public object Data;
            public RegistryValueKind Kind;
        }
        public List<Value> Values = new List<Value>();
        public Dictionary<string, RegistryImage> Children = new Dictionary<string, RegistryImage>();
        public static RegistryImage Read(RegistryKey key)
        {
            if (key == null) return null;
            var image = new RegistryImage();
            foreach (string name in key.GetValueNames()) image.Values.Add(new Value
                { Name = name, Data = key.GetValue(name, null, RegistryValueOptions.DoNotExpandEnvironmentNames), Kind = key.GetValueKind(name) });
            foreach (string child in key.GetSubKeyNames())
                using (RegistryKey nested = key.OpenSubKey(child)) image.Children.Add(child, Read(nested));
            return image;
        }
        public void Write(RegistryKey key)
        {
            foreach (Value value in Values) key.SetValue(value.Name, value.Data, value.Kind);
            foreach (var child in Children) using (RegistryKey nested = key.CreateSubKey(child.Key)) child.Value.Write(nested);
        }
    }

    internal sealed class InstallerEngine
    {
        public const string ClassId = "{B18A80F9-540D-4F9A-9F1D-A4798AC2A398}";
        public const string ProgId = "SlidePace.PowerPointAddIn";
        public const string ProductVersion = "1.0.2";
        public const string Marker = "slidepace-installation.txt";
        private readonly InstallOptions options;
        private readonly Action<string, int> progress;
        public string Root { get { return Path.GetFullPath(options.Root); } }
        internal string ClassKey { get { return Map(@"Software\Classes\CLSID\" + ClassId); } }
        internal string ProgIdKey { get { return Map(@"Software\Classes\" + ProgId); } }
        internal string AddInKey { get { return Map(@"Software\Microsoft\Office\PowerPoint\Addins\" + ProgId); } }
        internal string UninstallKey { get { return Map(@"Software\Microsoft\Windows\CurrentVersion\Uninstall\SlidePace"); } }
        public InstallerEngine(InstallOptions value, Action<string, int> onProgress) { options = value; progress = onProgress ?? delegate { }; }
        private string Map(string path) { return string.IsNullOrEmpty(options.RegistryPrefix) ? path : options.RegistryPrefix + "\\" + path; }
        private string[] Keys { get { return new[] { ClassKey, ProgIdKey, AddInKey, UninstallKey }; } }
        private void CheckOfficeClosed()
        {
            if (!options.TestMode && Process.GetProcessesByName("POWERPNT").Length > 0)
                throw new InvalidOperationException("PowerPoint 正在运行。请保存文件并关闭所有 PowerPoint 窗口，然后点击“重试”。");
        }
        private RegistryKey UserRegistry() { return RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64); }

        public string Install()
        {
            progress("正在检查 PowerPoint 和运行环境…", 5);
            if (!options.TestMode) OfficeEnvironment.Detect();
            CheckOfficeClosed();
            string versionDirectory = Path.Combine(Root, "v" + ProductVersion + "-" + Guid.NewGuid().ToString("N").Substring(0, 8));
            string uninstallPath = Path.Combine(Root, "Uninstall.exe");
            string uninstallBackup = uninstallPath + "." + Guid.NewGuid().ToString("N") + ".backup";
            bool rootExisted = Directory.Exists(Root);
            bool uninstallExisted = File.Exists(uninstallPath);
            string previousMarker = File.Exists(Path.Combine(Root, Marker)) ? File.ReadAllText(Path.Combine(Root, Marker)) : null;
            if (rootExisted && Directory.GetFileSystemEntries(Root).Length > 0 && previousMarker != ClassId)
                throw new InvalidOperationException("安装目录已存在其他文件，无法确认属于 SlidePace。请先检查该目录，再重新安装。");
            var images = new Dictionary<string, RegistryImage>();
            using (RegistryKey registry = UserRegistry())
                foreach (string path in Keys) using (RegistryKey key = registry.OpenSubKey(path)) images[path] = RegistryImage.Read(key);
            try
            {
                progress("正在部署计时插件…", 25);
                Directory.CreateDirectory(versionDirectory);
                string library = Path.Combine(versionDirectory, "SlidePace.AddIn.dll");
                using (Stream payload = options.PayloadOverrideForTests == null ? Assembly.GetExecutingAssembly().GetManifestResourceStream("SlidePace.Payload.dll") : new MemoryStream(options.PayloadOverrideForTests))
                {
                    if (payload == null) throw new InvalidOperationException("安装文件缺少插件，请重新获取完整安装包。");
                    using (var output = File.Create(library)) payload.CopyTo(output);
                }
                AssemblyName assembly = AssemblyName.GetAssemblyName(library);
                if (assembly.Name != "SlidePace.AddIn") throw new InvalidOperationException("插件文件校验失败。");
                File.WriteAllText(Path.Combine(versionDirectory, Marker), ClassId);
                if (File.Exists(uninstallPath)) File.Copy(uninstallPath, uninstallBackup, true);
                File.Copy(options.SetupPath, uninstallPath, true);
                progress("正在注册 PowerPoint 加载项…", 60);
                using (RegistryKey registry = UserRegistry())
                {
                    using (RegistryKey clsid = registry.CreateSubKey(ClassKey))
                    {
                        clsid.SetValue("", "SlidePace PowerPoint Add-in");
                        using (RegistryKey prog = clsid.CreateSubKey("ProgId")) prog.SetValue("", ProgId);
                        using (RegistryKey server = clsid.CreateSubKey("InprocServer32"))
                        {
                            server.SetValue("", "mscoree.dll");
                            server.SetValue("ThreadingModel", "Both");
                            WriteServerValues(server, assembly, library);
                            using (RegistryKey version = server.CreateSubKey(assembly.Version.ToString())) WriteServerValues(version, assembly, library);
                        }
                    }
                    using (RegistryKey prog = registry.CreateSubKey(ProgIdKey))
                    {
                        prog.SetValue("", "SlidePace PowerPoint Add-in");
                        using (RegistryKey clsid = prog.CreateSubKey("CLSID")) clsid.SetValue("", ClassId);
                    }
                    using (RegistryKey addin = registry.CreateSubKey(AddInKey))
                    {
                        addin.SetValue("FriendlyName", "SlidePace 计时");
                        addin.SetValue("Description", "顺计时、倒计时与系统时间；放映自动开始，双屏同步显示。");
                        addin.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                        addin.SetValue("CommandLineSafe", 0, RegistryValueKind.DWord);
                    }
                    using (RegistryKey uninstall = registry.CreateSubKey(UninstallKey))
                    {
                        uninstall.SetValue("DisplayName", "SlidePace PowerPoint 计时插件");
                        uninstall.SetValue("DisplayVersion", ProductVersion);
                        uninstall.SetValue("Publisher", "SlidePace");
                        uninstall.SetValue("InstallLocation", Root);
                        uninstall.SetValue("DisplayIcon", uninstallPath);
                        uninstall.SetValue("UninstallString", "\"" + uninstallPath + "\" --uninstall");
                        uninstall.SetValue("NoModify", 1, RegistryValueKind.DWord);
                        uninstall.SetValue("NoRepair", 1, RegistryValueKind.DWord);
                    }
                }
                if (options.FailAfterRegistryForTest) throw new IOException("Simulated failure after registration.");
                File.WriteAllText(Path.Combine(Root, Marker), ClassId);
                if (File.Exists(uninstallBackup)) File.Delete(uninstallBackup);
                // Registration points to the newly deployed version before older files are removed.
                foreach (string directory in Directory.GetDirectories(Root, "v*"))
                    if (!string.Equals(directory, versionDirectory, StringComparison.OrdinalIgnoreCase) && IsOwnedVersion(directory))
                        try { DeleteOwnedDirectory(directory); } catch (IOException) { } catch (UnauthorizedAccessException) { }
                progress("安装完成。重新打开 PowerPoint 即可使用。", 100);
                return library;
            }
            catch
            {
                using (RegistryKey registry = UserRegistry())
                {
                    foreach (var image in images)
                    {
                        registry.DeleteSubKeyTree(image.Key, false);
                        if (image.Value != null) using (RegistryKey key = registry.CreateSubKey(image.Key)) image.Value.Write(key);
                    }
                }
                if (Directory.Exists(versionDirectory)) DeleteOwnedDirectory(versionDirectory);
                if (File.Exists(uninstallBackup))
                {
                    File.Copy(uninstallBackup, uninstallPath, true);
                    File.Delete(uninstallBackup);
                }
                else if (!uninstallExisted && File.Exists(uninstallPath)) File.Delete(uninstallPath);
                if (previousMarker != null) File.WriteAllText(Path.Combine(Root, Marker), previousMarker);
                else if (File.Exists(Path.Combine(Root, Marker))) File.Delete(Path.Combine(Root, Marker));
                if (!rootExisted && Directory.Exists(Root) && Directory.GetFileSystemEntries(Root).Length == 0) Directory.Delete(Root);
                throw;
            }
        }
        private static void WriteServerValues(RegistryKey key, AssemblyName assembly, string library)
        {
            key.SetValue("Class", "SlidePace.PowerPointAddIn");
            key.SetValue("Assembly", assembly.FullName);
            key.SetValue("RuntimeVersion", "v4.0.30319");
            key.SetValue("CodeBase", new Uri(library).AbsoluteUri);
        }
        private void DeleteOwnedDirectory(string directory)
        {
            string full = Path.GetFullPath(directory);
            string prefix = Root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            if (!full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) throw new InvalidOperationException("拒绝清理安装目录之外的路径。");
            Directory.Delete(full, true);
        }
        private bool IsOwnedVersion(string directory)
        {
            string marker = Path.Combine(directory, Marker);
            return Path.GetFileName(directory).StartsWith("v", StringComparison.Ordinal) && File.Exists(marker) && File.ReadAllText(marker) == ClassId;
        }
        public void Uninstall(bool deleteSettings)
        {
            CheckOfficeClosed();
            progress("正在移除插件注册…", 20);
            using (RegistryKey registry = UserRegistry()) foreach (string key in Keys) registry.DeleteSubKeyTree(key, false);
            if (Directory.Exists(Root) && File.Exists(Path.Combine(Root, Marker)) && File.ReadAllText(Path.Combine(Root, Marker)) == ClassId)
            {
                foreach (string directory in Directory.GetDirectories(Root, "v*")) if (IsOwnedVersion(directory)) DeleteOwnedDirectory(directory);
                string uninstaller = Path.Combine(Root, "Uninstall.exe");
                if (File.Exists(uninstaller) && !string.Equals(Path.GetFullPath(uninstaller), Path.GetFullPath(options.SetupPath), StringComparison.OrdinalIgnoreCase)) File.Delete(uninstaller);
                File.Delete(Path.Combine(Root, Marker));
            }
            string settingsMarker = Path.Combine(options.SettingsRoot, "slidepace-settings-owner.txt");
            if (deleteSettings && File.Exists(settingsMarker) && File.ReadAllText(settingsMarker) == ClassId)
            {
                foreach (string name in new[] { "settings.json", "diagnostics.log", "slidepace-settings-owner.txt" })
                {
                    string file = Path.Combine(options.SettingsRoot, name);
                    if (File.Exists(file)) File.Delete(file);
                }
                if (Directory.GetFileSystemEntries(options.SettingsRoot).Length == 0) Directory.Delete(options.SettingsRoot);
            }
            progress("插件已卸载。", 100);
        }
    }
}
