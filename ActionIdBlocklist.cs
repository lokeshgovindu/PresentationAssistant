using System;
using System.Collections.Generic;
using System.Linq;

namespace PresentationAssistant
{
    /// <summary>
    /// Commands that are never announced. Split into the built-in list below, which
    /// suppresses commands the IDE fires constantly on its own, and a user list from the
    /// "Excluded Commands" setting.
    /// </summary>
    internal static class ActionIdBlocklist
    {
        private static readonly HashSet<string> ActionIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "Edit.BreakLine",
            "Edit.LineStart",
            "Edit.LineEnd",
            "Edit.LineUp",
            "Edit.LineDown",
            "Edit.CharLeft",
            "Edit.CharRight",
            "Edit.PageUp",
            "Edit.PageDown",

            // Backspace and Delete: bound, so Shortcuts Only does not filter them, but
            // announcing every one of them while typing is pure noise.
            "Edit.DeleteBackwards",
            "Edit.Delete",

            // Toolbar combo boxes fire on every selection change
            "Debug.DebugType",
            "Build.SolutionPlatforms",
            "Build.SolutionConfigurations",

            // This is coming always in Debug
            "Debug.LocationToolbar.ProcessCombo",
            "Debug.LocationToolbar.StackFrameCombo",
            "Debug.LocationToolbar.ThreadCombo"
        };

        /// <summary>The built-in names, for documentation and for the options page description.</summary>
        public static IEnumerable<string> BuiltIn => ActionIds.OrderBy(x => x, StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// True when the command is suppressed by the built-in list. Pass the canonical
        /// command name (<c>Edit.ScrollLineDown</c>), not the localized one, or these
        /// comparisons stop matching on a translated IDE.
        /// </summary>
        public static bool IsBlocked(string actionId)
        {
            if (string.IsNullOrEmpty(actionId)) return false;

            return ActionIds.Contains(actionId) ||
                   actionId.StartsWith("Debug.LocationToolbar.", StringComparison.OrdinalIgnoreCase);
        }
    }
}
