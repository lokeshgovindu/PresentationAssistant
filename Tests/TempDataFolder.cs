using Microsoft.VisualStudio.TestTools.UnitTesting;
using PresentationAssistant;
using System;
using System.IO;

namespace PresentationAssistant.Tests
{
    /// <summary>
    /// Redirects <see cref="AppPaths"/> at a throwaway folder for the duration of a test.
    /// Without this the theme tests would read and write the developer's real
    /// %APPDATA%\PresentationAssistant files.
    /// </summary>
    internal sealed class TempDataFolder : IDisposable
    {
        private readonly string _previous;

        public TempDataFolder()
        {
            _previous = AppPaths.DataFolderOverride;

            Path_ = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "PresentationAssistant.Tests",
                Guid.NewGuid().ToString("N"));

            Directory.CreateDirectory(Path_);
            AppPaths.DataFolderOverride = Path_;
        }

        public string Path_ { get; }

        public string ThemesFile => AppPaths.ThemesFile;

        public void WriteThemes(string json) => File.WriteAllText(ThemesFile, json);

        public string ReadThemes() => File.Exists(ThemesFile) ? File.ReadAllText(ThemesFile) : null;

        public void DeleteThemes()
        {
            if (File.Exists(ThemesFile)) File.Delete(ThemesFile);
        }

        public void Dispose()
        {
            AppPaths.DataFolderOverride = _previous;

            try { Directory.Delete(Path_, recursive: true); }
            catch (IOException) { /* a stray handle should not fail the test */ }
        }
    }

    /// <summary>Assertion helpers shared by the theme tests.</summary>
    internal static class AssertColour
    {
        public static void Hex(string expected, System.Windows.Media.Color actual, string because)
        {
            Assert.AreEqual(
                expected.ToUpperInvariant(),
                Theming.ThemePalette.ToHex(actual),
                because);
        }
    }
}
