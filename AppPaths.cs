using System;
using System.IO;

namespace PresentationAssistant
{
    /// <summary>
    /// Single place that knows where the extension keeps its per-user files.
    /// </summary>
    internal static class AppPaths
    {
        /// <summary>%APPDATA%\PresentationAssistant</summary>
        public static string DataFolder => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            PresentationAssistantPackage.ApplicationName);

        /// <summary>%APPDATA%\PresentationAssistant\presentationassistant.json</summary>
        public static string SettingsFile => Path.Combine(
            DataFolder,
            PresentationAssistantPackage.ApplicationName.ToLowerInvariant() + ".json");

        /// <summary>%APPDATA%\PresentationAssistant\themes.json</summary>
        public static string ThemesFile => Path.Combine(DataFolder, "themes.json");

        /// <summary>
        /// %APPDATA%\PresentationAssistant\themes.reference.json - a generated listing of
        /// the built-in themes to copy from. Never read back.
        /// </summary>
        public static string ThemesReferenceFile => Path.Combine(DataFolder, "themes.reference.json");

        public static void EnsureDataFolder() => Directory.CreateDirectory(DataFolder);
    }
}
