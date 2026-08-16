using System;
using System.Runtime.Serialization;
using System.Windows.Media;

namespace PresentationAssistant.Theming
{
    /// <summary>
    /// One entry in themes.json. Every field except <see cref="Name"/> is optional: a
    /// definition is merged onto a baseline (the built-in of the same name, or the
    /// default palette for a brand new theme), so a file can override just the one
    /// color it cares about.
    /// </summary>
    [DataContract]
    public class ThemeDefinition
    {
        [DataMember(Order = 0)]
        public string Name { get; set; }

        [DataMember(Order = 1)]
        public string Background { get; set; }

        [DataMember(Order = 2)]
        public string Foreground { get; set; }

        /// <summary>Color of the "via &lt;shortcut&gt;" and "×N" runs. Derived from the other two when omitted.</summary>
        [DataMember(Order = 3)]
        public string SecondaryForeground { get; set; }

        [DataMember(Order = 4)]
        public string Border { get; set; }

        [DataMember(Order = 5)]
        public double? Opacity { get; set; }

        /// <summary>
        /// Marks the theme as intended for a dark shell, which is what makes it eligible
        /// for the Auto pairing. Inferred from the background brightness when omitted.
        /// </summary>
        [DataMember(Order = 6)]
        public bool? IsDark { get; set; }

        /// <summary>
        /// Merges this definition onto <paramref name="baseline"/>. Unparseable colors
        /// fall back to the baseline value rather than failing the whole file.
        /// </summary>
        public ThemePalette ToPalette(ThemePalette baseline)
        {
            // Test whether a colour actually parsed, not merely whether the field was
            // present: a typo must not silently re-derive a colour the user never touched.
            var hasBackground = TryParseColor(Background, out var background);
            if (!hasBackground) background = baseline.BackgroundColor;

            var hasForeground = TryParseColor(Foreground, out var foreground);
            if (!hasForeground) foreground = baseline.ForegroundColor;

            // When the main colours are given, derive the secondary rather than
            // inheriting one that no longer suits them.
            var secondary = hasBackground || hasForeground
                ? ThemePalette.Blend(foreground, background, 0.45)
                : baseline.SecondaryForegroundColor;
            if (TryParseColor(SecondaryForeground, out var explicitSecondary)) secondary = explicitSecondary;

            if (!TryParseColor(Border, out var border)) border = baseline.BorderColor;

            return new ThemePalette(
                string.IsNullOrWhiteSpace(Name) ? baseline.Name : Name.Trim(),
                background,
                foreground,
                secondary,
                border,
                Clamp(Opacity ?? baseline.Opacity, 0.1, 1.0),
                IsDark ?? ThemePalette.Luminance(background) < 0.5);
        }

        private static double Clamp(double value, double min, double max)
        {
            return value < min ? min : (value > max ? max : value);
        }

        private static bool TryParseColor(string value, out Color color)
        {
            color = default(Color);
            if (string.IsNullOrWhiteSpace(value)) return false;

            try
            {
                var converted = ColorConverter.ConvertFromString(value.Trim());
                if (!(converted is Color parsed)) return false;

                color = parsed;
                return true;
            }
            catch (Exception)
            {
                // "#GGG", "chartreusey", etc.
                return false;
            }
        }
    }
}
