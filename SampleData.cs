namespace PresentationAssistant
{
    /// <summary>
    /// Design-time data for <c>PresentationAssistantWindow.xaml</c> so the XAML
    /// designer renders something representative.
    /// </summary>
    public static class SampleData
    {
        public static ShortcutDetails NoShortcut => new ShortcutDetails
        {
            ActionId    = "File.StartWindow",
            Description = "Start Window",
            Shortcuts   = null,
            Multiplier  = 1
        };

        public static ShortcutDetails MultiplierShortcut => new ShortcutDetails
        {
            ActionId    = "View.Output",
            Description = "Output",
            Shortcuts   = new[] { "Ctrl+Alt+O", "Alt+2" },
            Multiplier  = 9
        };
    }
}
