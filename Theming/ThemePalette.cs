using System;
using System.Windows.Media;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// One overlay look: four colors plus an opacity. Colors are kept alongside frozen
    /// brushes so a palette can both paint the window and act as the baseline that a
    /// partial <see cref="ThemeDefinition"/> from themes.json is merged onto.
    /// </summary>
    public sealed class ThemePalette
    {
        /// <summary>Resource keys the window and its XAML agree on.</summary>
        public const string BackgroundKey          = "PA.BackgroundBrush";
        public const string ForegroundKey          = "PA.ForegroundBrush";
        public const string SecondaryForegroundKey = "PA.SecondaryForegroundBrush";
        public const string BorderKey              = "PA.BorderBrush";

        public ThemePalette(
            string name,
            Color background,
            Color foreground,
            Color secondaryForeground,
            Color border,
            double opacity,
            bool isDark)
        {
            Name                     = name;
            BackgroundColor          = background;
            ForegroundColor          = foreground;
            SecondaryForegroundColor = secondaryForeground;
            BorderColor              = border;
            Opacity                  = opacity;
            IsDark                   = isDark;

            Background          = Freeze(background);
            Foreground          = Freeze(foreground);
            SecondaryForeground = Freeze(secondaryForeground);
            Border              = Freeze(border);
        }

        public string Name { get; }

        /// <summary>
        /// Whether this palette is meant for a dark shell. Used to pick the pair that
        /// the Auto theme switches between.
        /// </summary>
        public bool IsDark { get; }

        public double Opacity { get; }

        public Color BackgroundColor { get; }

        /// <summary>Color of the bold command description.</summary>
        public Color ForegroundColor { get; }

        /// <summary>Color of the dimmer "via &lt;shortcut&gt;" and "×N" runs.</summary>
        public Color SecondaryForegroundColor { get; }

        public Color BorderColor { get; }

        public Brush Background { get; }

        public Brush Foreground { get; }

        public Brush SecondaryForeground { get; }

        public Brush Border { get; }

        /// <summary>
        /// Builds the palette that matches the running IDE, from three colors read off
        /// the shell. Kept separate from the shell lookup itself so it can be exercised
        /// against known Visual Studio palettes without a live shell.
        /// </summary>
        /// <param name="background">The shell's tool window background.</param>
        /// <param name="foreground">The shell's tool window text color.</param>
        /// <param name="accent">The shell's accent, used for the border and as the tint.</param>
        public static ThemePalette FromShellColors(
            string name,
            Color background,
            Color foreground,
            Color accent,
            double opacity)
        {
            // Using the tool window background verbatim makes the overlay vanish into the
            // IDE, which defeats the point. Nudging it towards the accent keeps it
            // recognisably "Visual Studio" while still reading as a callout.
            var tinted = Blend(background, accent, 0.14);

            // Guarantee the description stays legible even if the shell hands us a
            // foreground that is close to the tinted background.
            var text = Contrast(foreground, tinted) < 3.0
                ? (Luminance(tinted) < 0.5 ? Colors.White : Colors.Black)
                : foreground;

            return new ThemePalette(
                name,
                tinted,
                text,
                Blend(text, tinted, 0.42),
                accent,
                opacity,
                Luminance(tinted) < 0.5);
        }

        /// <summary>
        /// WCAG contrast ratio between two colors, in the 1..21 range. Used only to catch
        /// a shell palette that would render the overlay unreadable.
        /// </summary>
        public static double Contrast(Color a, Color b)
        {
            var la = RelativeLuminance(a);
            var lb = RelativeLuminance(b);
            var lighter = Math.Max(la, lb);
            var darker = Math.Min(la, lb);
            return (lighter + 0.05) / (darker + 0.05);
        }

        private static double RelativeLuminance(Color c)
        {
            return 0.2126 * Linear(c.R) + 0.7152 * Linear(c.G) + 0.0722 * Linear(c.B);
        }

        private static double Linear(byte channel)
        {
            var v = channel / 255.0;
            return v <= 0.03928 ? v / 12.92 : Math.Pow((v + 0.055) / 1.055, 2.4);
        }

        /// <summary>Returns the same palette under a different name.</summary>
        public ThemePalette Rename(string name)
        {
            return new ThemePalette(name, BackgroundColor, ForegroundColor,
                SecondaryForegroundColor, BorderColor, Opacity, IsDark);
        }

        /// <summary>Mixes <paramref name="amount"/> of <paramref name="towards"/> into <paramref name="from"/>.</summary>
        public static Color Blend(Color from, Color towards, double amount)
        {
            return Color.FromRgb(
                Mix(from.R, towards.R, amount),
                Mix(from.G, towards.G, amount),
                Mix(from.B, towards.B, amount));
        }

        /// <summary>Perceived brightness in the 0..1 range, used to tell dark shells from light ones.</summary>
        public static double Luminance(Color color)
        {
            return (0.299 * color.R + 0.587 * color.G + 0.114 * color.B) / 255.0;
        }

        private static byte Mix(byte a, byte b, double amount) => (byte)(a + (b - a) * amount);

        private static Brush Freeze(Color color)
        {
            var brush = new SolidColorBrush(color);
            brush.Freeze();
            return brush;
        }
    }
}
