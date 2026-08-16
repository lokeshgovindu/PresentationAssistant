using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using PresentationAssistant.State;
using PresentationAssistant.Theming;
using System.IO;

namespace PresentationAssistant.Tests
{
    [TestClass]
    public class SettingsTests
    {
        [TestMethod]
        public void A_non_positive_window_timeout_is_clamped()
        {
            // Left alone, Task.Delay throws on a negative TimeSpan and the exception escapes
            // onto a discarded task - which used to leave the overlay stuck on screen.
            var settings = new Settings { WindowTimeoutInMS = -1 }.Normalized();

            Assert.IsTrue(settings.WindowTimeoutInMS >= 250, "actual: " + settings.WindowTimeoutInMS);
        }

        [TestMethod]
        public void Absurd_timeouts_are_clamped_at_both_ends()
        {
            Assert.AreEqual(600000, new Settings { WindowTimeoutInMS = int.MaxValue }.Normalized().WindowTimeoutInMS);
            Assert.AreEqual(600000, new Settings { MultiplierTimeoutInMS = int.MaxValue }.Normalized().MultiplierTimeoutInMS);
            Assert.AreEqual(250, new Settings { MultiplierTimeoutInMS = 0 }.Normalized().MultiplierTimeoutInMS);
        }

        [TestMethod]
        public void Font_size_is_clamped()
        {
            Assert.AreEqual(8, new Settings { FontSize = 0 }.Normalized().FontSize);
            Assert.AreEqual(96, new Settings { FontSize = 1000 }.Normalized().FontSize);
            Assert.AreEqual(24, new Settings { FontSize = 24 }.Normalized().FontSize);
        }

        [TestMethod]
        public void A_missing_theme_falls_back_to_the_default()
        {
            Assert.AreEqual(ThemeNames.Default, new Settings { Theme = null }.Normalized().Theme);
            Assert.AreEqual(ThemeNames.Default, new Settings { Theme = "  " }.Normalized().Theme);
        }

        [TestMethod]
        public void A_missing_exclusion_list_becomes_empty_rather_than_null()
        {
            Assert.AreEqual(0, new Settings { ExcludedCommands = null }.Normalized().ExcludedCommands.Length);
        }

        [TestMethod]
        public void Defaults_are_what_a_fresh_install_gets()
        {
            var settings = new Settings();

            Assert.IsTrue(settings.Enabled, "the overlay is on out of the box");
            Assert.AreEqual(ThemeNames.VisualStudio, settings.Theme, "matches the running IDE by default");
            Assert.AreEqual(5000, settings.WindowTimeoutInMS);
            Assert.AreEqual(10000, settings.MultiplierTimeoutInMS);
            Assert.AreEqual(24, settings.FontSize);
            Assert.AreEqual(OverlayLayout.Horizontal, settings.Layout);
            Assert.IsFalse(settings.ShortcutsOnly);
            Assert.IsFalse(settings.Diagnostics);
        }

        [TestMethod]
        public void The_file_round_trips_every_setting()
        {
            using (var data = new TempDataFolder())
            {
                var path = Path.Combine(data.Path_, "round-trip.json");

                var original = new Settings
                {
                    Enabled = false,
                    WindowTimeoutInMS = 4000,
                    MultiplierTimeoutInMS = 8000,
                    ShortcutsOnly = true,
                    Theme = "Nord",
                    FontSize = 32,
                    Layout = OverlayLayout.Vertical,
                    Diagnostics = true,
                    ExcludedCommands = new[] { "Edit.Line*", "View.Output" },
                };

                original.SaveToFile(path);
                var reloaded = JsonConvert.DeserializeObject<Settings>(File.ReadAllText(path));

                Assert.IsFalse(reloaded.Enabled);
                Assert.AreEqual(4000, reloaded.WindowTimeoutInMS);
                Assert.AreEqual(8000, reloaded.MultiplierTimeoutInMS);
                Assert.IsTrue(reloaded.ShortcutsOnly);
                Assert.AreEqual("Nord", reloaded.Theme);
                Assert.AreEqual(32, reloaded.FontSize);
                Assert.AreEqual(OverlayLayout.Vertical, reloaded.Layout);
                Assert.IsTrue(reloaded.Diagnostics);
                CollectionAssert.AreEqual(new[] { "Edit.Line*", "View.Output" }, reloaded.ExcludedCommands);
            }
        }

        [TestMethod]
        public void Enums_are_written_as_names_so_the_file_stays_readable()
        {
            using (var data = new TempDataFolder())
            {
                var path = Path.Combine(data.Path_, "readable.json");
                new Settings { Layout = OverlayLayout.Vertical }.SaveToFile(path);

                var json = File.ReadAllText(path);

                StringAssert.Contains(json, "\"Vertical\"", "Layout should not be a bare number");
                StringAssert.Contains(json, "\"Theme\": \"VisualStudio\"");
            }
        }

        [TestMethod]
        public void An_older_file_without_the_newer_keys_still_loads()
        {
            // Exactly the shape shipped before theming existed.
            var json = @"{ ""WindowTimeoutInMS"": 3000, ""MultiplierTimeoutInMS"": 7000, ""ShortcutsOnly"": true }";

            var settings = JsonConvert.DeserializeObject<Settings>(json).Normalized();

            Assert.AreEqual(3000, settings.WindowTimeoutInMS, "existing values are kept");
            Assert.IsTrue(settings.ShortcutsOnly);
            Assert.AreEqual(ThemeNames.Default, settings.Theme, "the new setting takes its default");
            Assert.IsTrue(settings.Enabled);
            Assert.AreEqual(24, settings.FontSize);
        }
    }
}
