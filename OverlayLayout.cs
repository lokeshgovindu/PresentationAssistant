namespace PresentationAssistant
{
    /// <summary>How the overlay arranges the command name and its shortcut.</summary>
    public enum OverlayLayout
    {
        /// <summary>One line: <c>Scroll Line Down via Ctrl+Down Arrow x9</c>.</summary>
        Horizontal = 0,

        /// <summary>
        /// Two lines, the command above its shortcut. Keeps the overlay narrow when
        /// command names are long.
        /// </summary>
        Vertical
    }
}
