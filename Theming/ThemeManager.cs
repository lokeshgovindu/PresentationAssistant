using Microsoft.VisualStudio.PlatformUI;
using Microsoft.VisualStudio.Shell;
using System;
using System.Windows.Media;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// Turns a theme name into a concrete <see cref="ThemePalette"/> and raises
    /// <see cref="ThemeChanged"/> when the shell theme is switched, so the Auto and
    /// VisualStudio selections can follow along.
    /// </summary>
    internal static class ThemeManager
    {
        // Resolve runs once per announced command, and for the shell-derived themes it
        // would otherwise re-query shell colors and rebuild brushes every time. Cache the
        // result and invalidate on the only three things that can change it.
        private static ThemePalette _cached;
        private static string _cachedName;
        private static int _cachedCatalogVersion = -1;
        private static int _cachedShellGeneration = -1;
        private static int _shellGeneration;

        static ThemeManager()
        {
            // Fires on the UI thread whenever the user switches the Visual Studio theme.
            VSColorTheme.ThemeChanged += OnVsColorThemeChanged;
        }

        /// <summary>Raised after the shell theme changed; palettes should be re-resolved.</summary>
        public static event EventHandler ThemeChanged;

        /// <summary>
        /// Resolves <paramref name="name"/> against the catalog. Must be called on the UI
        /// thread, because Auto and VisualStudio query shell colors. An unknown name -
        /// a theme deleted from themes.json, say - falls back to the Auto pair rather than
        /// failing.
        /// </summary>
        public static ThemePalette Resolve(string name)
        {
            var catalog = ThemeCatalog.Current;

            if (string.IsNullOrWhiteSpace(name)) name = ThemeNames.Auto;
            name = name.Trim();

            if (_cached != null &&
                _cachedCatalogVersion == ThemeCatalog.Version &&
                _cachedShellGeneration == _shellGeneration &&
                string.Equals(_cachedName, name, StringComparison.OrdinalIgnoreCase))
            {
                return _cached;
            }

            var resolved = ResolveCore(catalog, name);

            _cached = resolved;
            _cachedName = name;
            _cachedCatalogVersion = ThemeCatalog.Version;
            _cachedShellGeneration = _shellGeneration;

            return resolved;
        }

        private static ThemePalette ResolveCore(ThemeCatalog catalog, string name)
        {
            if (string.Equals(name, ThemeNames.VisualStudio, StringComparison.OrdinalIgnoreCase))
            {
                return catalog.ApplyDynamicOverride(ThemeNames.VisualStudio, FromShell());
            }

            if (string.Equals(name, ThemeNames.Auto, StringComparison.OrdinalIgnoreCase))
            {
                var chosen = IsShellDark() ? catalog.DarkDefault : catalog.LightDefault;
                return catalog.ApplyDynamicOverride(ThemeNames.Auto, chosen);
            }

            return catalog.TryGet(name, out var palette)
                ? palette
                : (IsShellDark() ? catalog.DarkDefault : catalog.LightDefault);
        }

        /// <summary>
        /// True when the active Visual Studio theme is a dark one. Decided from the
        /// brightness of the tool window background rather than from a theme GUID, so
        /// third-party and custom themes are classified correctly too.
        /// </summary>
        public static bool IsShellDark()
        {
            try
            {
                var background = ShellColor(EnvironmentColors.ToolWindowBackgroundColorKey, Colors.White);
                return ThemePalette.Luminance(background) < 0.5;
            }
            catch (Exception)
            {
                // No shell available (XAML designer, unit tests) - assume light.
                return false;
            }
        }

        private static ThemePalette FromShell()
        {
            try
            {
                var background = ShellColor(EnvironmentColors.ToolWindowBackgroundColorKey,
                    ThemeCatalog.Fallback.BackgroundColor);
                var foreground = ShellColor(EnvironmentColors.ToolWindowTextColorKey,
                    ThemeCatalog.Fallback.ForegroundColor);

                // The accent drives both the border and the background tint. Fall back
                // through a couple of keys, since which ones a theme defines varies.
                var accent = ShellColor(EnvironmentColors.SystemHighlightColorKey, Colors.Transparent);
                if (accent == Colors.Transparent)
                {
                    accent = ShellColor(EnvironmentColors.AccentBorderColorKey,
                        ThemeCatalog.Fallback.BorderColor);
                }

                return ThemePalette.FromShellColors(
                    ThemeNames.VisualStudio, background, foreground, accent, 0.92);
            }
            catch (Exception)
            {
                return ThemeCatalog.Fallback;
            }
        }

        private static Color ShellColor(ThemeResourceKey key, Color fallback)
        {
            var color = VSColorTheme.GetThemedColor(key);

            // A fully transparent color means the key was not themed; don't build a
            // palette out of it. Otherwise drop the alpha - a partly transparent brush
            // would compound with the window's own Opacity.
            return color.A == 0
                ? fallback
                : Color.FromRgb(color.R, color.G, color.B);
        }

        private static void OnVsColorThemeChanged(ThemeChangedEventArgs e)
        {
            _shellGeneration++;
            ThemeChanged?.Invoke(null, EventArgs.Empty);
        }
    }
}
