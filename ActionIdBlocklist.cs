using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PresentationAssistant
{
    internal static class ActionIdBlocklist
    {
        private static readonly HashSet<string> ActionIds = new HashSet<string> {
            "Edit.BreakLine",
            "Edit.LineStart",
            "Edit.LineEnd",
            "Edit.LineUp",
            "Edit.LineDown",
            "Edit.CharLeft",
            "Edit.CharRight",
            "Edit.PageUp",
            "Edit.PageDown",

            // This is coming always in Debug
            "Debug.LocationToolbar.ProcessCombo",
            "Debug.LocationToolbar.StackFrameCombo",
            "Debug.LocationToolbar.ThreadCombo"
        };

        public static bool IsBlocked(string actionId)
        {
            return
                ActionIds.Contains(actionId) ||
                actionId.StartsWith("Debug.LocationToolbar.");
        }
    }
}
