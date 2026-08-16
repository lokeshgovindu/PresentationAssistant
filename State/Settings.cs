using Newtonsoft.Json;
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

        /// <summary>Raised after <see cref="Save"/> so the running package can pick the changes up.</summary>
        public static event EventHandler SettingsUpdated;

        public static Settings Load()
        {
            try
            {
                AppPaths.EnsureDataFolder();
                return JsonConvert.DeserializeObject<Settings>(File.ReadAllText(GetSettingsFilePath()))
                       ?? new Settings();
            }
            catch (Exception ex)
            {
                // A corrupt or unreadable file must not take the package down with it.
                Debug.WriteLine($"[{PresentationAssistantPackage.ApplicationName}] Failed to load settings: {ex}");
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
                Debug.WriteLine($"[{PresentationAssistantPackage.ApplicationName}] Failed to save settings: {ex}");
            }

            OnSettingsUpdated(this, EventArgs.Empty);
        }

        public void SaveToFile(string path)
        {
            File.WriteAllText(path, JsonConvert.SerializeObject(this, Formatting.Indented));
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
