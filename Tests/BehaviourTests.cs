using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using System.Threading;

namespace PresentationAssistant.Tests
{
    [TestClass]
    public class CommandExclusionsTests
    {
        private static CommandExclusions From(string patterns) =>
            new CommandExclusions(CommandExclusions.Split(patterns));

        [TestMethod]
        public void Exact_and_prefix_patterns_match()
        {
            var exclusions = From("Edit.Line*; View.Output");

            Assert.IsTrue(exclusions.IsExcluded("Edit.LineDown"), "prefix");
            Assert.IsTrue(exclusions.IsExcluded("edit.linedown"), "case insensitive");
            Assert.IsTrue(exclusions.IsExcluded("View.Output"), "exact");

            Assert.IsFalse(exclusions.IsExcluded("View.OutputWindow"), "exact must not match a longer name");
            Assert.IsFalse(exclusions.IsExcluded("Edit.Copy"));
            Assert.IsFalse(exclusions.IsExcluded(""));
            Assert.IsFalse(exclusions.IsExcluded(null));
        }

        [TestMethod]
        public void A_bare_star_is_ignored_rather_than_silencing_everything()
        {
            var exclusions = From("*");

            Assert.IsTrue(exclusions.IsEmpty);
            Assert.IsFalse(exclusions.IsExcluded("Edit.Copy"));
        }

        [TestMethod]
        public void Split_accepts_semicolons_commas_and_newlines()
        {
            var parsed = CommandExclusions.Split(" Edit.A ;;\r\n Edit.B , Edit.C ");

            CollectionAssert.AreEqual(new[] { "Edit.A", "Edit.B", "Edit.C" }, parsed);
        }

        [TestMethod]
        public void Split_of_nothing_is_empty()
        {
            Assert.AreEqual(0, CommandExclusions.Split(null).Length);
            Assert.AreEqual(0, CommandExclusions.Split("   ").Length);
        }

        [TestMethod]
        public void Join_round_trips_through_Split()
        {
            var joined = CommandExclusions.Join(new[] { "Edit.A", "Edit.B" });

            Assert.AreEqual("Edit.A; Edit.B", joined);
            Assert.AreEqual(2, CommandExclusions.Split(joined).Length);
        }

        [TestMethod]
        public void Empty_excludes_nothing()
        {
            Assert.IsFalse(CommandExclusions.Empty.IsExcluded("Edit.Copy"));
            Assert.IsTrue(CommandExclusions.Empty.IsEmpty);
        }
    }

    [TestClass]
    public class ActionIdBlocklistTests
    {
        [TestMethod]
        public void The_noisy_editor_commands_are_blocked()
        {
            foreach (var blocked in new[]
            {
                "Edit.LineUp", "Edit.CharLeft", "Edit.PageDown", "Edit.BreakLine",
                "Build.SolutionConfigurations", "Debug.DebugType",
            })
            {
                Assert.IsTrue(ActionIdBlocklist.IsBlocked(blocked), blocked + " should be blocked");
            }
        }

        [TestMethod]
        public void Backspace_and_delete_are_blocked()
        {
            // Both are bound, so Shortcuts Only does not filter them; announcing every one
            // while typing is noise. Raised in a Marketplace review.
            Assert.IsTrue(ActionIdBlocklist.IsBlocked("Edit.DeleteBackwards"));
            Assert.IsTrue(ActionIdBlocklist.IsBlocked("Edit.Delete"));
        }

        [TestMethod]
        public void The_debugger_location_toolbar_is_blocked_by_prefix()
        {
            Assert.IsTrue(ActionIdBlocklist.IsBlocked("Debug.LocationToolbar.ProcessCombo"));
            Assert.IsTrue(ActionIdBlocklist.IsBlocked("Debug.LocationToolbar.SomethingNew"));
        }

        [TestMethod]
        public void Matching_ignores_case_and_tolerates_nothing()
        {
            Assert.IsTrue(ActionIdBlocklist.IsBlocked("edit.lineup"));
            Assert.IsFalse(ActionIdBlocklist.IsBlocked("Edit.ScrollLineDown"));
            Assert.IsFalse(ActionIdBlocklist.IsBlocked(null));
            Assert.IsFalse(ActionIdBlocklist.IsBlocked(""));
        }

        [TestMethod]
        public void The_built_in_list_is_exposed_for_documentation()
        {
            Assert.IsTrue(ActionIdBlocklist.BuiltIn.Any());
        }
    }

    [TestClass]
    public class ShortcutDisplayStatisticsTests
    {
        [TestMethod]
        public void Repeats_of_the_same_command_accumulate()
        {
            var statistics = new ShortcutDisplayStatistics(10000);

            statistics.OnAction("Edit.ScrollLineDown");
            Assert.AreEqual(1, statistics.Multiplier, "the first press is not a repeat");

            statistics.OnAction("Edit.ScrollLineDown");
            statistics.OnAction("Edit.ScrollLineDown");
            Assert.AreEqual(3, statistics.Multiplier);
        }

        [TestMethod]
        public void A_different_command_starts_a_new_run()
        {
            var statistics = new ShortcutDisplayStatistics(10000);

            statistics.OnAction("Edit.ScrollLineDown");
            statistics.OnAction("Edit.ScrollLineDown");
            statistics.OnAction("View.Output");

            Assert.AreEqual(1, statistics.Multiplier);
            Assert.AreEqual("View.Output", statistics.LastActionId);
        }

        [TestMethod]
        public void A_gap_longer_than_the_timeout_starts_a_new_run()
        {
            var statistics = new ShortcutDisplayStatistics(1);

            statistics.OnAction("Edit.ScrollLineDown");
            Thread.Sleep(30);
            statistics.OnAction("Edit.ScrollLineDown");

            Assert.AreEqual(1, statistics.Multiplier, "the run should have expired");
        }

        [TestMethod]
        public void The_timeout_can_be_changed_at_runtime()
        {
            var statistics = new ShortcutDisplayStatistics(1000);
            statistics.SetMultiplierTimeout(60000);

            Assert.AreEqual(60000, statistics.MultiplierTimeout.TotalMilliseconds, 0.001);
        }
    }

    [TestClass]
    public class ShortcutDetailsTests
    {
        [TestMethod]
        public void Several_bindings_are_joined_for_display()
        {
            var details = new ShortcutDetails { Shortcuts = new[] { "Ctrl+Alt+O", "Alt+2" } };

            Assert.IsTrue(details.HasShortcuts);
            Assert.AreEqual("Ctrl+Alt+O or Alt+2", details.ShortcutsStr);
        }

        [TestMethod]
        public void No_bindings_reads_as_no_shortcuts()
        {
            Assert.IsFalse(new ShortcutDetails { Shortcuts = null }.HasShortcuts);
            Assert.IsFalse(new ShortcutDetails { Shortcuts = new string[0] }.HasShortcuts);
            Assert.AreEqual(string.Empty, new ShortcutDetails { Shortcuts = null }.ShortcutsStr);
        }

        [TestMethod]
        public void The_repeat_count_only_shows_from_two_upwards()
        {
            Assert.IsFalse(new ShortcutDetails { Multiplier = 1 }.HasMultiplier);
            Assert.AreEqual(string.Empty, new ShortcutDetails { Multiplier = 1 }.MultiplierStr);

            Assert.IsTrue(new ShortcutDetails { Multiplier = 9 }.HasMultiplier);
            Assert.AreEqual("×9", new ShortcutDetails { Multiplier = 9 }.MultiplierStr);
        }

        [TestMethod]
        public void Sample_data_is_usable_by_the_designer_and_the_preview()
        {
            Assert.IsTrue(SampleData.MultiplierShortcut.HasShortcuts);
            Assert.IsTrue(SampleData.MultiplierShortcut.HasMultiplier);
            Assert.IsFalse(SampleData.NoShortcut.HasShortcuts);
        }
    }
}
