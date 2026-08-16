using System;
using System.Linq;

namespace PresentationAssistant
{
    /// <summary>
    /// Everything the overlay needs to know about a single executed command.
    /// This is the view model bound to <see cref="PresentationAssistantWindow"/>.
    /// </summary>
    public class ShortcutDetails
    {
        /// <summary>Canonical command name, e.g. <c>Edit.ScrollLineDown</c>.</summary>
        public string ActionId { get; set; }

        /// <summary>Human readable form of <see cref="ActionId"/>, e.g. "Scroll Line Down".</summary>
        public string Description { get; set; }

        /// <summary>
        /// How many times in a row this command has been invoked inside the multiplier
        /// window. 1 means "just once", and is not rendered.
        /// </summary>
        public int Multiplier { get; set; } = 1;

        /// <summary>Key bindings for the command, or <c>null</c> when it has none.</summary>
        public string[] Shortcuts { get; set; }

        public bool HasShortcuts => Shortcuts != null && Shortcuts.Length > 0;

        public bool HasMultiplier => Multiplier > 1;

        public string ShortcutsStr => HasShortcuts ? string.Join(" or ", Shortcuts) : string.Empty;

        /// <summary>The repeat count as it is displayed, e.g. "×9". Empty for a single press.</summary>
        public string MultiplierStr => HasMultiplier ? "×" + Multiplier : string.Empty;

        public override string ToString()
        {
            return $"ActionId: [{ActionId}], Description: [{Description}], " +
                   $"Shortcuts: [{ShortcutsStr}], Multiplier: [{Multiplier}]";
        }
    }
}
