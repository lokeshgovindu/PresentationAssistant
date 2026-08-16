namespace PresentationAssistant.Theming
{
    /// <summary>
    /// Theme names with behaviour attached. Everything else is just a palette name,
    /// looked up in <see cref="ThemeCatalog"/> - which is what lets themes.json add new
    /// ones without a code change.
    /// </summary>
    public static class ThemeNames
    {
        /// <summary>
        /// Follow Visual Studio: <see cref="Classic"/> under the Blue/Light themes,
        /// <see cref="ClassicDark"/> under Dark. Re-evaluated when the shell theme changes.
        /// </summary>
        public const string Auto = "Auto";

        /// <summary>Built from the shell's own colors, so the overlay matches the running IDE.</summary>
        public const string VisualStudio = "VisualStudio";

        /// <summary>
        /// The original pale green overlay the extension shipped with, and the light half
        /// of <see cref="Auto"/>.
        /// </summary>
        public const string Classic = "Classic";

        /// <summary>Dark counterpart of <see cref="Classic"/>, and the dark half of <see cref="Auto"/>.</summary>
        public const string ClassicDark = "ClassicDark";

        /// <summary>Earlier builds called <see cref="Classic"/> "Light"; keep it resolving.</summary>
        public const string LegacyLight = "Light";

        /// <summary>Earlier builds called <see cref="ClassicDark"/> "Dark"; keep it resolving.</summary>
        public const string LegacyDark = "Dark";

        /// <summary>
        /// What a fresh install gets: match the running IDE. <see cref="Auto"/> is the
        /// alternative if you would rather ship a fixed pair of palettes that only
        /// switches on light/dark.
        /// </summary>
        public const string Default = VisualStudio;
    }
}
