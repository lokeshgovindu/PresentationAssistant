using Microsoft.VisualStudio.TestTools.UnitTesting;
using PresentationAssistant.Theming;
using System.Windows.Media;

namespace PresentationAssistant.Tests
{
    [TestClass]
    public class ThemeDefinitionTests
    {
        private static ThemePalette Baseline() => new ThemePalette(
            "Baseline",
            background:          Hex("#DCF2DC"),
            foreground:          Hex("#1E1E1E"),
            secondaryForeground: Hex("#5A6B5A"),
            border:              Hex("#7FA97F"),
            opacity:             0.85,
            isDark:              false);

        [TestMethod]
        public void Omitted_fields_are_inherited()
        {
            var palette = new ThemeDefinition { Name = "Tweaked", Opacity = 1.0 }.ToPalette(Baseline());

            Assert.AreEqual("Tweaked", palette.Name);
            Assert.AreEqual(1.0, palette.Opacity, 0.001);
            AssertColour.Hex("#DCF2DC", palette.BackgroundColor, "background inherited");
            AssertColour.Hex("#5A6B5A", palette.SecondaryForegroundColor, "secondary inherited");
        }

        [TestMethod]
        public void Secondary_is_derived_when_the_main_colours_change()
        {
            var palette = new ThemeDefinition
            {
                Name = "Dark thing",
                Background = "#101820",
                Foreground = "#F2F2F2",
            }.ToPalette(Baseline());

            AssertColour.Hex("#101820", palette.BackgroundColor, "background applied");
            Assert.AreNotEqual("#5A6B5A", ThemePalette.ToHex(palette.SecondaryForegroundColor),
                "inheriting the old secondary would not suit the new colours");
        }

        [TestMethod]
        public void An_explicit_secondary_wins_over_the_derived_one()
        {
            var palette = new ThemeDefinition
            {
                Name = "Explicit",
                Background = "#101820",
                Foreground = "#F2F2F2",
                SecondaryForeground = "#8AA0B0",
            }.ToPalette(Baseline());

            AssertColour.Hex("#8AA0B0", palette.SecondaryForegroundColor, "explicit value used");
        }

        [TestMethod]
        public void An_unparseable_colour_changes_nothing_else()
        {
            var palette = new ThemeDefinition { Name = "Typo", Background = "#NOTACOLOUR" }
                .ToPalette(Baseline());

            AssertColour.Hex("#DCF2DC", palette.BackgroundColor, "falls back to the baseline");
            AssertColour.Hex("#5A6B5A", palette.SecondaryForegroundColor,
                "a typo must not silently re-derive a colour the user never touched");
        }

        [TestMethod]
        public void Opacity_is_clamped_into_a_usable_range()
        {
            Assert.AreEqual(1.0, new ThemeDefinition { Name = "A", Opacity = 5.0 }.ToPalette(Baseline()).Opacity, 0.001);
            Assert.AreEqual(0.1, new ThemeDefinition { Name = "B", Opacity = -3.0 }.ToPalette(Baseline()).Opacity, 0.001);
        }

        [TestMethod]
        public void IsDark_is_inferred_from_the_background_when_omitted()
        {
            Assert.IsTrue(new ThemeDefinition { Name = "D", Background = "#101820" }
                .ToPalette(Baseline()).IsDark);

            Assert.IsFalse(new ThemeDefinition { Name = "L", Background = "#F5F5F5" }
                .ToPalette(Baseline()).IsDark);
        }

        [TestMethod]
        public void An_explicit_IsDark_is_respected_even_against_the_brightness()
        {
            Assert.IsTrue(new ThemeDefinition { Name = "Odd", Background = "#F5F5F5", IsDark = true }
                .ToPalette(Baseline()).IsDark);
        }

        [TestMethod]
        public void A_blank_name_keeps_the_baseline_name()
        {
            Assert.AreEqual("Baseline",
                new ThemeDefinition { Name = "   ", Opacity = 0.5 }.ToPalette(Baseline()).Name);
        }

        private static Color Hex(string value) => (Color)ColorConverter.ConvertFromString(value);
    }
}
