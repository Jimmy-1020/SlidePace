using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

namespace SlidePace
{
    public enum OverlayPosition { TopRight, TopLeft, BottomRight, BottomLeft, Custom }

    [DataContract]
    public sealed class SavedLocation
    {
        [DataMember] public string Device;
        [DataMember] public double X;
        [DataMember] public double Y;
    }

    [DataContract]
    public sealed class UserSettings
    {
        [DataMember] public int Mode;
        [DataMember] public int CountdownSeconds = 600;
        [DataMember] public CountdownEndBehavior CountdownEndBehavior;
        [DataMember] public string OvertimeColor = "#FF0000";
        [DataMember] public int FontSize = 1;
        [DataMember] public string NumberFontName = "Consolas";
        [DataMember] public bool ShowPresenterTimer = true;
        [DataMember] public OverlayPosition PresenterPosition;
        [DataMember] public OverlayPosition AudiencePosition;
        [DataMember] public string PresenterDevice = "";
        [DataMember] public string AudienceDevice = "";
        [DataMember] public List<SavedLocation> Locations = new List<SavedLocation>();

        [OnDeserializing]
        private void SetDefaults(StreamingContext context)
        {
            CountdownSeconds = 600;
            CountdownEndBehavior = SlidePace.CountdownEndBehavior.ContinueCountUp;
            OvertimeColor = "#FF0000";
            FontSize = 1;
            NumberFontName = "Consolas";
            ShowPresenterTimer = true;
            PresenterDevice = AudienceDevice = "";
            Locations = new List<SavedLocation>();
        }

        public Color GetOvertimeColor()
        {
            try { return ColorTranslator.FromHtml(OvertimeColor); }
            catch { return Color.Red; }
        }

        public void Validate()
        {
            if (!Enum.IsDefined(typeof(TimerMode), Mode)) Mode = 0;
            if (CountdownSeconds < 1 || CountdownSeconds > 86399) CountdownSeconds = 600;
            if (!Enum.IsDefined(typeof(CountdownEndBehavior), CountdownEndBehavior)) CountdownEndBehavior = SlidePace.CountdownEndBehavior.ContinueCountUp;
            if (FontSize < 0 || FontSize > 2) FontSize = 1;
            try
            {
                using (var family = new FontFamily(NumberFontName ?? ""))
                {
                    if (!family.IsStyleAvailable(FontStyle.Regular) && !family.IsStyleAvailable(FontStyle.Bold))
                        throw new ArgumentException();
                    NumberFontName = family.Name;
                }
            }
            catch { NumberFontName = "Consolas"; }
            if (!Enum.IsDefined(typeof(OverlayPosition), PresenterPosition)) PresenterPosition = OverlayPosition.TopRight;
            if (!Enum.IsDefined(typeof(OverlayPosition), AudiencePosition)) AudiencePosition = OverlayPosition.TopRight;
            try
            {
                if (string.IsNullOrWhiteSpace(OvertimeColor)) throw new FormatException();
                Color color = ColorTranslator.FromHtml(OvertimeColor);
                if (color.A != 255) throw new FormatException();
                OvertimeColor = ColorTranslator.ToHtml(color);
            }
            catch { OvertimeColor = "#FF0000"; }
            PresenterDevice = PresenterDevice ?? "";
            AudienceDevice = AudienceDevice ?? "";
            if (Locations == null) Locations = new List<SavedLocation>();
            Locations.RemoveAll(delegate(SavedLocation item)
            {
                return item == null || string.IsNullOrEmpty(item.Device) || double.IsNaN(item.X) || double.IsNaN(item.Y) ||
                    double.IsInfinity(item.X) || double.IsInfinity(item.Y);
            });
            foreach (SavedLocation location in Locations)
            {
                location.X = Math.Max(0, Math.Min(1, location.X));
                location.Y = Math.Max(0, Math.Min(1, location.Y));
            }
        }
    }

    public sealed class SettingsStore
    {
        public const string OwnershipFile = "slidepace-settings-owner.txt";
        public const string OwnershipValue = "{B18A80F9-540D-4F9A-9F1D-A4798AC2A398}";
        public string DirectoryPath { get; private set; }
        public string FilePath { get { return Path.Combine(DirectoryPath, "settings.json"); } }
        public SettingsStore(string directory) { DirectoryPath = directory; }
        public static string DefaultDirectory
        {
            get { return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SlidePace"); }
        }
        public UserSettings Load()
        {
            try
            {
                using (var stream = File.OpenRead(FilePath))
                {
                    var settings = (UserSettings)new DataContractJsonSerializer(typeof(UserSettings)).ReadObject(stream);
                    if (settings == null) throw new SerializationException();
                    settings.Validate();
                    // A PowerPoint session always starts with all timer modes deselected.
                    // Older settings may still contain the previous session's selection.
                    settings.Mode = (int)TimerMode.None;
                    return settings;
                }
            }
            catch { return new UserSettings(); }
        }
        public void Save(UserSettings settings)
        {
            settings.Validate();
            Directory.CreateDirectory(DirectoryPath);
            File.WriteAllText(Path.Combine(DirectoryPath, OwnershipFile), OwnershipValue);
            string temporary = FilePath + "." + Guid.NewGuid().ToString("N") + ".tmp";
            try
            {
                using (var stream = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                {
                    new DataContractJsonSerializer(typeof(UserSettings)).WriteObject(stream, settings);
                    stream.Flush(true);
                }
                if (File.Exists(FilePath)) File.Replace(temporary, FilePath, null);
                else File.Move(temporary, FilePath);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
        public void Log(Exception error)
        {
            try
            {
                Directory.CreateDirectory(DirectoryPath);
                File.WriteAllText(Path.Combine(DirectoryPath, OwnershipFile), OwnershipValue);
                string path = Path.Combine(DirectoryPath, "diagnostics.log");
                if (File.Exists(path) && new FileInfo(path).Length > 1024 * 1024) File.Delete(path);
                File.AppendAllText(path, DateTime.Now.ToString("s") + " " + error + Environment.NewLine);
            }
            catch { }
        }
    }
}
