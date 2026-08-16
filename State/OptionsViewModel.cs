using PresentationAssistant.Theming;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Controls;
using System.Windows.Media;

namespace PresentationAssistant.State
{
    /// <summary>
    /// Backs the options page. Holds the editable settings and derives the live preview
    /// from them, so changing a theme, size or layout is visible before applying.
    /// </summary>
    internal sealed class OptionsViewModel : INotifyPropertyChanged
    {
        private bool _enabled = true;
        private bool _shortcutsOnly;
        private bool _diagnostics;
        private int _windowTimeout = 5000;
        private int _multiplierTimeout = 10000;
        private int _fontSize = Settings.DefaultFontSize;
        private string _theme = ThemeNames.Default;
        private OverlayLayout _layout = OverlayLayout.Horizontal;
        private string _excludedCommands = string.Empty;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool Enabled
        {
            get => _enabled;
            set => Set(ref _enabled, value);
        }

        public bool ShortcutsOnly
        {
            get => _shortcutsOnly;
            set => Set(ref _shortcutsOnly, value);
        }

        public bool Diagnostics
        {
            get => _diagnostics;
            set => Set(ref _diagnostics, value);
        }

        public int WindowTimeout
        {
            get => _windowTimeout;
            set => Set(ref _windowTimeout, value);
        }

        public int MultiplierTimeout
        {
            get => _multiplierTimeout;
            set => Set(ref _multiplierTimeout, value);
        }

        public int FontSize
        {
            get => _fontSize;
            set { if (Set(ref _fontSize, value)) RaisePreviewChanged(); }
        }

        public string Theme
        {
            get => _theme;
            set
            {
                if (!Set(ref _theme, value)) return;

                // Switching theme discards half-made colour edits and shows the new
                // theme's own colours, which is the only reading that is not surprising.
                _coloursEdited = false;
                LoadColoursFromTheme();
            }
        }

        public OverlayLayout Layout
        {
            get => _layout;
            set
            {
                if (!Set(ref _layout, value)) return;

                Raise(nameof(IsHorizontal));
                Raise(nameof(IsVertical));
                RaisePreviewChanged();
            }
        }

        /// <summary>Radio button binding for <see cref="OverlayLayout.Horizontal"/>.</summary>
        public bool IsHorizontal
        {
            get => Layout == OverlayLayout.Horizontal;
            set { if (value) Layout = OverlayLayout.Horizontal; }
        }

        /// <summary>Radio button binding for <see cref="OverlayLayout.Vertical"/>.</summary>
        public bool IsVertical
        {
            get => Layout == OverlayLayout.Vertical;
            set { if (value) Layout = OverlayLayout.Vertical; }
        }

        /// <summary>
        /// The exclusion patterns as free text. Accepts one per line as well as the
        /// semicolon separated form the settings file uses.
        /// </summary>
        public string ExcludedCommands
        {
            get => _excludedCommands;
            set => Set(ref _excludedCommands, value);
        }

        public IReadOnlyList<string> AvailableThemes { get; private set; } = new string[0];

        public string ThemesFilePath => AppPaths.ThemesFile;

        #region Colours

        private string _backgroundHex = string.Empty;
        private string _foregroundHex = string.Empty;
        private string _secondaryHex = string.Empty;
        private string _borderHex = string.Empty;
        private double _colourOpacity = 1.0;
        private bool _coloursEdited;

        /// <summary>
        /// The four colours as hex text. Held as strings rather than Colors so a
        /// half-typed value in the text box is harmless, and so themes.json can be written
        /// straight from them.
        /// </summary>
        public string BackgroundHex
        {
            get => _backgroundHex;
            set { if (Set(ref _backgroundHex, value)) OnColourEdited(); }
        }

        public string ForegroundHex
        {
            get => _foregroundHex;
            set { if (Set(ref _foregroundHex, value)) OnColourEdited(); }
        }

        public string SecondaryHex
        {
            get => _secondaryHex;
            set { if (Set(ref _secondaryHex, value)) OnColourEdited(); }
        }

        public string BorderHex
        {
            get => _borderHex;
            set { if (Set(ref _borderHex, value)) OnColourEdited(); }
        }

        public double ColourOpacity
        {
            get => _colourOpacity;
            set { if (Set(ref _colourOpacity, Math.Round(value, 2))) OnColourEdited(); }
        }

        /// <summary>True when themes.json already carries an entry for the selected theme.</summary>
        public bool HasCustomColours { get; private set; }

        /// <summary>Reverts the selected theme to its built-in colours.</summary>
        public void ResetColours()
        {
            ThemeCatalog.RemoveOverride(Theme);
            _coloursEdited = false;
            LoadColoursFromTheme();
        }

        /// <summary>
        /// Persists edited colours as a themes.json entry for the selected theme. Only
        /// writes when something was actually changed, so simply opening the page and
        /// pressing OK does not start overriding built-ins.
        /// </summary>
        public void SaveColoursIfEdited()
        {
            if (!_coloursEdited) return;

            ThemeCatalog.SaveOverride(new ThemeDefinition
            {
                Name                = Theme,
                Background          = BackgroundHex,
                Foreground          = ForegroundHex,
                SecondaryForeground = SecondaryHex,
                Border              = BorderHex,
                Opacity             = ColourOpacity,
            });

            _coloursEdited = false;
        }

        private void OnColourEdited()
        {
            _coloursEdited = true;
            RaisePreviewChanged();
        }

        private void LoadColoursFromTheme()
        {
            var palette = Palette;

            _backgroundHex = ThemePalette.ToHex(palette.BackgroundColor);
            _foregroundHex = ThemePalette.ToHex(palette.ForegroundColor);
            _secondaryHex  = ThemePalette.ToHex(palette.SecondaryForegroundColor);
            _borderHex     = ThemePalette.ToHex(palette.BorderColor);
            _colourOpacity = Math.Round(palette.Opacity, 2);

            HasCustomColours = ThemeCatalog.HasOverride(Theme);

            foreach (var name in new[]
            {
                nameof(BackgroundHex), nameof(ForegroundHex), nameof(SecondaryHex),
                nameof(BorderHex), nameof(ColourOpacity), nameof(HasCustomColours),
            })
            {
                Raise(name);
            }

            RaisePreviewChanged();
        }

        #endregion

        #region Live preview

        public ShortcutDetails PreviewShortcut => SampleData.MultiplierShortcut;

        // Driven by the edited colours, not the stored palette, so the preview is what the
        // overlay will actually look like once applied.
        public Brush PreviewBackground => BrushFrom(BackgroundHex, Palette.Background);

        public Brush PreviewForeground => BrushFrom(ForegroundHex, Palette.Foreground);

        public Brush PreviewSecondaryForeground => BrushFrom(SecondaryHex, Palette.SecondaryForeground);

        public Brush PreviewBorder => BrushFrom(BorderHex, Palette.Border);

        public double PreviewOpacity => ColourOpacity;

        private static Brush BrushFrom(string hex, Brush fallback)
        {
            if (!ThemePalette.TryParse(hex, out var colour)) return fallback;

            var brush = new SolidColorBrush(colour);
            brush.Freeze();
            return brush;
        }

        public Orientation PreviewOrientation =>
            Layout == OverlayLayout.Vertical ? Orientation.Vertical : Orientation.Horizontal;

        /// <summary>The word "via" only belongs in the single line layout.</summary>
        public System.Windows.Visibility PreviewViaVisibility =>
            Layout == OverlayLayout.Horizontal
                ? System.Windows.Visibility.Visible
                : System.Windows.Visibility.Collapsed;

        /// <summary>Spacing between the command and its shortcut: beside it, or below it.</summary>
        public System.Windows.Thickness PreviewDetailMargin =>
            Layout == OverlayLayout.Horizontal
                ? new System.Windows.Thickness(8, 0, 0, 0)
                : new System.Windows.Thickness(0, 2, 0, 0);

        private ThemePalette Palette => ThemeManager.Resolve(Theme);

        private void RaisePreviewChanged()
        {
            Raise(nameof(PreviewBackground));
            Raise(nameof(PreviewForeground));
            Raise(nameof(PreviewSecondaryForeground));
            Raise(nameof(PreviewBorder));
            Raise(nameof(PreviewOpacity));
            Raise(nameof(PreviewOrientation));
            Raise(nameof(PreviewViaVisibility));
            Raise(nameof(PreviewDetailMargin));
        }

        #endregion

        public void Load(Settings settings)
        {
            // Re-read the catalog so themes added to themes.json appear straight away.
            ThemeCatalog.Invalidate();
            AvailableThemes = ThemeCatalog.Current.Names.ToList();
            Raise(nameof(AvailableThemes));

            _enabled           = settings.Enabled;
            _shortcutsOnly     = settings.ShortcutsOnly;
            _diagnostics       = settings.Diagnostics;
            _windowTimeout     = settings.WindowTimeoutInMS;
            _multiplierTimeout = settings.MultiplierTimeoutInMS;
            _fontSize          = settings.FontSize;
            _theme             = settings.Theme;
            _layout            = settings.Layout;

            // One per line reads far better than a single long line.
            _excludedCommands = string.Join("\r\n", settings.ExcludedCommands ?? new string[0]);

            foreach (var name in new[]
            {
                nameof(Enabled), nameof(ShortcutsOnly), nameof(Diagnostics), nameof(WindowTimeout),
                nameof(MultiplierTimeout), nameof(FontSize), nameof(Theme), nameof(Layout),
                nameof(IsHorizontal), nameof(IsVertical), nameof(ExcludedCommands),
            })
            {
                Raise(name);
            }

            _coloursEdited = false;
            LoadColoursFromTheme();
        }

        public Settings ToSettings()
        {
            return new Settings
            {
                Enabled               = Enabled,
                ShortcutsOnly         = ShortcutsOnly,
                Diagnostics           = Diagnostics,
                WindowTimeoutInMS     = WindowTimeout,
                MultiplierTimeoutInMS = MultiplierTimeout,
                FontSize              = FontSize,
                Theme                 = Theme,
                Layout                = Layout,
                ExcludedCommands      = CommandExclusions.Split(ExcludedCommands),
            }.Normalized();
        }

        private bool Set<T>(ref T field, T value, [CallerMemberName] string property = null)
        {
            if (EqualityComparer<T>.Default.Equals(field, value)) return false;

            field = value;
            Raise(property);
            return true;
        }

        private void Raise(string property)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(property));
        }
    }
}
