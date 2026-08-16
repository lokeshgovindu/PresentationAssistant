using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using PresentationAssistant.Theming;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.Serialization;

namespace PresentationAssistant.State
{
    /// <summary>
    /// User settings, persisted as JSON next to the other per-user application data so
    /// they survive Visual Studio upgrades and can be edited by hand.
    /// </summary>
    [DataContract]
    public class Settings
    {
        public const int DefaultFontSize = 24;

        // Guard rails for the numeric settings. Both the options grid and a hand-edited
        // file can supply anything, and a non-positive window timeout used to leave the
        // overlay stuck on screen: Task.Delay throws on a negative TimeSpan, and that
        // exception escaped onto a discarded task where nothing observed it.
        private const int MinTimeoutInMS = 250;
        private const int MaxTimeoutInMS = 600000;
        private const int MinFontSize = 8;
        private const int MaxFontSize = 96;

        /// <summary>How long the overlay stays on screen after the last command.</summary>
        [DataMember(Order = 0)]
        public int WindowTimeoutInMS { get; set; } = 5000;

        /// <summary>
        /// Maximum gap between two invocations of the same command for them to still
        /// count towards the same "×N" run.
        /// </summary>
        [DataMember(Order = 1)]
        public int MultiplierTimeoutInMS { get; set; } = 10000;

        /// <summary>When set, commands without a key binding are not announced.</summary>
        [DataMember(Order = 2)]
        public bool ShortcutsOnly { get; set; }

        /// <summary>
        /// Name of the overlay theme: "Auto", "VisualStudio", a built-in palette, or
        /// anything defined in themes.json. A plain string rather than an enum precisely
        /// so that user-defined themes are first-class.
        /// </summary>
        [DataMember(Order = 3)]
        public string Theme { get; set; } = ThemeNames.Default;

        /// <summary>
        /// Commands the user does not want announced, on top of the built-in blocklist.
        /// Matched against the canonical command name; a trailing <c>*</c> matches a
        /// prefix, e.g. <c>Edit.Line*</c>.
        /// </summary>
        [DataMember(Order = 4)]
        public string[] ExcludedCommands { get; set; } = new string[0];

        /// <summary>
        /// Master switch. When false nothing is announced, so the extension can be
        /// silenced without uninstalling it.
        /// </summary>
        [DataMember(Order = 5)]
        public bool Enabled { get; set; } = true;

        /// <summary>Overlay text size, in device independent units.</summary>
        [DataMember(Order = 6)]
        public int FontSize { get; set; } = DefaultFontSize;

        /// <summary>Whether the command and its shortcut sit on one line or two.</summary>
        [DataMember(Order = 7)]
        [JsonConverter(typeof(StringEnumConverter))]
        public OverlayLayout Layout { get; set; } = OverlayLayout.Horizontal;

        /// <summary>
        /// Writes what the extension is doing to a PresentationAssistant pane in the
        /// Output window. Off by default; useful when reporting a problem.
        /// </summary>
        [DataMember(Order = 8)]
        public bool Diagnostics { get; set; }

        /// <summary>Raised after <see cref="Save"/> so the running package can pick the changes up.</summary>
        public static event EventHandler SettingsUpdated;

        public static Settings Load()
        {
            try
            {
                AppPaths.EnsureDataFolder();
                var loaded = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(GetSettingsFilePath()))
                             ?? new Settings();
                return loaded.Normalized();
            }
            catch (Exception ex)
            {
                // A corrupt or unreadable file must not take the package down with it.
                Debug.WriteLine($"[{Product.Name}] Failed to load settings: {ex}");
                return new Settings();
            }
        }

        public void Save()
        {
            try
            {
                AppPaths.EnsureDataFolder();
                SaveToFile(GetSettingsFilePath());
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{Product.Name}] Failed to save settings: {ex}");
            }

            OnSettingsUpdated(this, EventArgs.Empty);
        }

        public void SaveToFile(string path)
        {
            File.WriteAllText(path, JsonConvert.SerializeObject(Normalized(), Formatting.Indented));
        }

        /// <summary>
        /// Brings every value into a range the rest of the extension can rely on, so
        /// neither the options grid nor a hand-edited file can put it into a broken state.
        /// </summary>
        public Settings Normalized()
        {
            WindowTimeoutInMS     = Clamp(WindowTimeoutInMS, MinTimeoutInMS, MaxTimeoutInMS);
            MultiplierTimeoutInMS = Clamp(MultiplierTimeoutInMS, MinTimeoutInMS, MaxTimeoutInMS);
            FontSize              = Clamp(FontSize, MinFontSize, MaxFontSize);

            if (string.IsNullOrWhiteSpace(Theme)) Theme = ThemeNames.Default;
            if (ExcludedCommands == null) ExcludedCommands = new string[0];

            return this;
        }

        private static int Clamp(int value, int min, int max)
        {
            return value < min ? min : (value > max ? max : value);
        }

        private static string GetSettingsFilePath()
        {
            var path = AppPaths.SettingsFile;

            if (!File.Exists(path) || string.IsNullOrWhiteSpace(File.ReadAllText(path)))
            {
                new Settings().SaveToFile(path);
            }

            return path;
        }

        private static void OnSettingsUpdated(object sender, EventArgs ea)
        {
            SettingsUpdated?.Invoke(sender, ea);
        }
    }
}
