using Microsoft.VisualStudio.TestTools.UnitTesting;
using PresentationAssistant.Theming;
using System.Windows.Media;

namespace PresentationAssistant.Tests
{
    [TestClass]
    public class ThemePaletteTests
    {
        [TestMethod]
        public void Hex_round_trips()
        {
            Assert.IsTrue(ThemePalette.TryParse("#DCF2DC", out var colour));
            Assert.AreEqual("#DCF2DC", ThemePalette.ToHex(colour));
        }

        [TestMethod]
        public void TryParse_accepts_what_themes_json_documents()
        {
            Assert.IsTrue(ThemePalette.TryParse("#102030", out _), "#RRGGBB");
            Assert.IsTrue(ThemePalette.TryParse("#80102030", out _), "#AARRGGBB");
            Assert.IsTrue(ThemePalette.TryParse("Gainsboro", out _), "named colour");
            Assert.IsTrue(ThemePalette.TryParse("  #102030  ", out _), "surrounding space");
        }

        [TestMethod]
        public void TryParse_refuses_nonsense_without_throwing()
        {
            Assert.IsFalse(ThemePalette.TryParse("#NOTACOLOUR", out _));
            Assert.IsFalse(ThemePalette.TryParse("chartreusey", out _));
            Assert.IsFalse(ThemePalette.TryParse("", out _));
            Assert.IsFalse(ThemePalette.TryParse(null, out _));
        }

        [TestMethod]
        public void Luminance_orders_black_below_white()
        {
            Assert.IsTrue(ThemePalette.Luminance(Colors.Black) < 0.5);
            Assert.IsTrue(ThemePalette.Luminance(Colors.White) > 0.5);
        }

        [TestMethod]
        public void Blend_moves_towards_the_target()
        {
            var midway = ThemePalette.Blend(Colors.Black, Colors.White, 0.5);

            Assert.AreEqual(127, midway.R, 1);
            Assert.AreEqual(Colors.Black, ThemePalette.Blend(Colors.Black, Colors.White, 0.0));
            Assert.AreEqual(Colors.White, ThemePalette.Blend(Colors.Black, Colors.White, 1.0));
        }

        [TestMethod]
        public void Contrast_is_highest_between_black_and_white()
        {
            Assert.AreEqual(21.0, ThemePalette.Contrast(Colors.Black, Colors.White), 0.1);
            Assert.AreEqual(1.0, ThemePalette.Contrast(Colors.Black, Colors.Black), 0.01);
        }

        [TestMethod]
        public void Shell_derivation_tints_away_from_the_raw_background()
        {
            var background = Hex("#252526");   // Visual Studio Dark tool window
            var palette = ThemePalette.FromShellColors("VS", background, Hex("#F1F1F1"), Hex("#0078D4"), 0.92);

            Assert.AreNotEqual(
                ThemePalette.ToHex(background),
                ThemePalette.ToHex(palette.BackgroundColor),
                "using the shell colour verbatim would make the overlay vanish into the chrome");

            Assert.IsTrue(palette.IsDark, "a dark shell yields a dark palette");
            Assert.IsTrue(ThemePalette.Contrast(palette.ForegroundColor, palette.BackgroundColor) >= 4.5,
                "the command name must clear WCAG AA");
        }

        [TestMethod]
        public void Shell_derivation_rescues_an_unreadable_foreground()
        {
            // A theme whose text colour is nearly its background: the overlay would be
            // illegible if the given foreground were used as-is.
            var palette = ThemePalette.FromShellColors("VS", Hex("#303030"), Hex("#333333"), Hex("#C05000"), 1.0);

            Assert.IsTrue(ThemePalette.Contrast(palette.ForegroundColor, palette.BackgroundColor) >= 4.5,
                "the contrast guard should have substituted black or white");
        }

        [TestMethod]
        public void Rename_keeps_every_colour_and_the_dark_flag()
        {
            var original = ThemePalette.FromShellColors("VS", Hex("#252526"), Hex("#F1F1F1"), Hex("#0078D4"), 0.9);
            var renamed = original.Rename("Something else");

            Assert.AreEqual("Something else", renamed.Name);
            Assert.AreEqual(original.BackgroundColor, renamed.BackgroundColor);
            Assert.AreEqual(original.ForegroundColor, renamed.ForegroundColor);
            Assert.AreEqual(original.SecondaryForegroundColor, renamed.SecondaryForegroundColor);
            Assert.AreEqual(original.BorderColor, renamed.BorderColor);
            Assert.AreEqual(original.Opacity, renamed.Opacity, 0.001);
            Assert.AreEqual(original.IsDark, renamed.IsDark);
        }

        private static Color Hex(string value) => (Color)ColorConverter.ConvertFromString(value);
    }
}
