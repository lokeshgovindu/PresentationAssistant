using Microsoft.VisualStudio.TestTools.UnitTesting;
using PresentationAssistant.Theming;
using System.Linq;

namespace PresentationAssistant.Tests
{
    [TestClass]
    public class ThemeCatalogTests
    {
        private TempDataFolder _data;

        [TestInitialize]
        public void SetUp()
        {
            _data = new TempDataFolder();
            ThemeCatalog.Invalidate();
        }

        [TestCleanup]
        public void TearDown()
        {
            _data.Dispose();
            ThemeCatalog.Invalidate();
        }

        [TestMethod]
        public void Dropdown_lists_the_two_computed_themes_first()
        {
            var names = ThemeCatalog.Current.Names;

            Assert.AreEqual(ThemeNames.Auto, names[0]);
            Assert.AreEqual(ThemeNames.VisualStudio, names[1]);
        }

        [TestMethod]
        public void Built_in_names_are_unique_and_include_the_recognisable_palettes()
        {
            var names = ThemeCatalog.Current.Names.ToList();

            CollectionAssert.AreEqual(
                names.Distinct(System.StringComparer.OrdinalIgnoreCase).ToList(),
                names,
                "the dropdown must not contain duplicates");

            foreach (var expected in new[] { "Classic", "ClassicDark", "Nord", "Dracula", "TokyoNight", "HighContrast" })
            {
                Assert.IsTrue(names.Contains(expected), expected + " is missing");
            }
        }

        [TestMethod]
        public void The_seeded_themes_file_is_created_and_adds_nothing()
        {
            // Automatic seeding is deliberately attempted only once per process, so that
            // deleting the file on purpose keeps it deleted. Go through the entry point the
            // options page uses, which creates it on demand.
            _data.DeleteThemes();
            var before = ThemeCatalog.Current.Names.Count;

            ThemeCatalog.EnsureAuthoringFiles();

            Assert.IsNotNull(_data.ReadThemes(), "themes.json should have been created");
            StringAssert.Contains(_data.ReadThemes(), "themes.reference.json",
                "the seeded file should point at the generated reference listing");

            ThemeCatalog.Invalidate();
            Assert.AreEqual(before, ThemeCatalog.Current.Names.Count,
                "the seeded file must parse and contribute no themes");
        }

        [TestMethod]
        public void The_generated_reference_lists_every_built_in_in_full()
        {
            ThemeCatalog.EnsureAuthoringFiles();

            var reference = System.IO.File.ReadAllText(AppPaths.ThemesReferenceFile);

            StringAssert.Contains(reference, "GENERATED", "it must warn that editing it does nothing");
            foreach (var name in new[] { "Classic", "ClassicDark", "Nord", "HighContrast" })
            {
                StringAssert.Contains(reference, "\"Name\": \"" + name + "\"");
            }

            StringAssert.Contains(reference, "\"Background\": \"#DCF2DC\"", "colours are listed in full");
        }

        [TestMethod]
        public void An_entry_overrides_a_built_in_in_place()
        {
            var baseline = ThemeCatalog.Current;
            var position = baseline.Names.ToList().IndexOf("Amber");

            _data.WriteThemes(@"[ { ""Name"": ""Amber"", ""Opacity"": 1.0 } ]");
            ThemeCatalog.Invalidate();

            var catalog = ThemeCatalog.Current;

            Assert.AreEqual(baseline.Names.Count, catalog.Names.Count, "override must not add an entry");
            Assert.AreEqual(position, catalog.Names.ToList().IndexOf("Amber"), "position must be kept");

            Assert.IsTrue(catalog.TryGet("Amber", out var amber));
            Assert.AreEqual(1.0, amber.Opacity, 0.001);
            AssertColour.Hex("#FDF1D6", amber.BackgroundColor, "unspecified fields are inherited");
        }

        [TestMethod]
        public void A_new_name_adds_a_theme_baselined_on_the_default()
        {
            _data.WriteThemes(@"[ { ""Name"": ""My Talk"", ""Background"": ""#101820"", ""IsDark"": true } ]");
            ThemeCatalog.Invalidate();

            Assert.IsTrue(ThemeCatalog.Current.TryGet("My Talk", out var mine));
            AssertColour.Hex("#101820", mine.BackgroundColor, "background comes from the file");
            Assert.IsTrue(mine.IsDark);
        }

        [TestMethod]
        public void Comments_are_tolerated_and_nameless_entries_ignored()
        {
            var before = ThemeCatalog.Current.Names.Count;

            _data.WriteThemes("// a comment\n[ { \"Background\": \"#123456\" } ]");
            ThemeCatalog.Invalidate();

            Assert.AreEqual(before, ThemeCatalog.Current.Names.Count,
                "an entry with no Name must be skipped");
        }

        [TestMethod]
        public void A_malformed_file_leaves_the_built_ins_working()
        {
            var before = ThemeCatalog.Current.Names.Count;

            _data.WriteThemes("{ this is not valid json");
            ThemeCatalog.Invalidate();

            Assert.AreEqual(before, ThemeCatalog.Current.Names.Count);
        }

        [TestMethod]
        public void Lookup_ignores_case_and_surrounding_space_and_legacy_names()
        {
            var catalog = ThemeCatalog.Current;

            Assert.IsTrue(catalog.TryGet("nORd", out _), "case insensitive");
            Assert.IsTrue(catalog.TryGet("  Nord  ", out _), "trimmed");
            Assert.IsFalse(catalog.TryGet("NoSuchTheme", out _));

            // Earlier builds called these Light and Dark.
            Assert.IsTrue(catalog.TryGet("Light", out var light));
            Assert.AreEqual(ThemeNames.Classic, light.Name);
            Assert.IsTrue(catalog.TryGet("Dark", out var dark));
            Assert.AreEqual(ThemeNames.ClassicDark, dark.Name);
        }

        [TestMethod]
        public void Auto_uses_the_classic_pair_and_follows_an_override_of_it()
        {
            AssertColour.Hex("#DCF2DC", ThemeCatalog.Current.LightDefault.BackgroundColor, "light half");
            AssertColour.Hex("#16241B", ThemeCatalog.Current.DarkDefault.BackgroundColor, "dark half");

            // Overriding by the legacy name must still reach the pair.
            _data.WriteThemes(@"[ { ""Name"": ""Dark"", ""Background"": ""#3B0D0D"" } ]");
            ThemeCatalog.Invalidate();

            AssertColour.Hex("#3B0D0D", ThemeCatalog.Current.DarkDefault.BackgroundColor,
                "Auto's dark half picks up the override");
        }

        [TestMethod]
        public void The_computed_themes_are_tunable_without_being_duplicated()
        {
            _data.WriteThemes(@"[
                { ""Name"": ""VisualStudio"", ""Opacity"": 0.55, ""Border"": ""#FF0000"" },
                { ""Name"": ""Auto"",         ""Opacity"": 0.40 }
            ]");
            ThemeCatalog.Invalidate();

            var catalog = ThemeCatalog.Current;

            Assert.AreEqual(1, catalog.Names.Count(n => n == ThemeNames.VisualStudio));
            Assert.AreEqual(1, catalog.Names.Count(n => n == ThemeNames.Auto));
            Assert.IsFalse(catalog.TryGet(ThemeNames.VisualStudio, out _),
                "VisualStudio stays computed rather than becoming a stored palette");

            Assert.IsTrue(catalog.TryGet("Nord", out var stand_in));
            var tuned = catalog.ApplyDynamicOverride(ThemeNames.VisualStudio, stand_in);

            Assert.AreEqual(0.55, tuned.Opacity, 0.001);
            AssertColour.Hex("#FF0000", tuned.BorderColor, "override reaches the border");
            AssertColour.Hex(
                ThemePalette.ToHex(stand_in.BackgroundColor), tuned.BackgroundColor,
                "unspecified fields keep the computed value");
            Assert.AreEqual(stand_in.Name, tuned.Name, "the computed palette keeps its name");
        }

        [TestMethod]
        public void A_name_with_no_override_passes_straight_through()
        {
            var catalog = ThemeCatalog.Current;
            Assert.IsTrue(catalog.TryGet("Nord", out var nord));

            Assert.AreSame(nord, catalog.ApplyDynamicOverride("Classic", nord));
        }

        [TestMethod]
        public void Saving_and_removing_an_override_round_trips()
        {
            Assert.IsFalse(ThemeCatalog.HasOverride("Nord"));

            ThemeCatalog.SaveOverride(new ThemeDefinition { Name = "Nord", Background = "#010203" });

            Assert.IsTrue(ThemeCatalog.HasOverride("Nord"));
            Assert.IsTrue(ThemeCatalog.Current.TryGet("Nord", out var edited));
            AssertColour.Hex("#010203", edited.BackgroundColor, "the saved colour is used");

            ThemeCatalog.RemoveOverride("Nord");

            Assert.IsFalse(ThemeCatalog.HasOverride("Nord"));
            Assert.IsTrue(ThemeCatalog.Current.TryGet("Nord", out var reverted));
            AssertColour.Hex("#2E3440", reverted.BackgroundColor, "back to the built-in");
        }

        [TestMethod]
        public void Saving_an_override_replaces_rather_than_appends()
        {
            ThemeCatalog.SaveOverride(new ThemeDefinition { Name = "Nord", Background = "#010203" });
            ThemeCatalog.SaveOverride(new ThemeDefinition { Name = "Nord", Background = "#040506" });

            var occurrences = ThemeCatalog.Current.Names.Count(n => n == "Nord");
            Assert.AreEqual(1, occurrences);

            Assert.IsTrue(ThemeCatalog.Current.TryGet("Nord", out var nord));
            AssertColour.Hex("#040506", nord.BackgroundColor, "the later save wins");
        }

        [TestMethod]
        public void Version_changes_when_the_catalog_is_rebuilt()
        {
            var first = ThemeCatalog.Version;
            var unchanged = ThemeCatalog.Version;
            Assert.AreEqual(first, unchanged, "reading twice must not bump the version");

            ThemeCatalog.Invalidate();
            var _ = ThemeCatalog.Current;

            Assert.AreNotEqual(first, ThemeCatalog.Version, "a rebuild must be observable");
        }
    }
}
