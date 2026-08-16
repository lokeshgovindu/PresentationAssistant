using System;
using System.IO;

namespace PresentationAssistant
{
    /// <summary>
    /// Single place that knows where the extension keeps its per-user files.
    /// </summary>
    internal static class AppPaths
    {
        /// <summary>
        /// Redirects every path below somewhere else. Only for tests, which must not read
        /// or write the real per-user files.
        /// </summary>
        internal static string DataFolderOverride { get; set; }

        /// <summary>%APPDATA%\PresentationAssistant</summary>
        public static string DataFolder => DataFolderOverride ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            Product.Name);

        /// <summary>%APPDATA%\PresentationAssistant\presentationassistant.json</summary>
        public static string SettingsFile => Path.Combine(
            DataFolder,
            Product.Name.ToLowerInvariant() + ".json");

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
