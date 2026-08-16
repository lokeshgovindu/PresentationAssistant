using PresentationAssistant.Theming;

namespace PresentationAssistant
{
    /// <summary>
    /// Everything about how the overlay should look for one showing. Grouped so the
    /// window has a single entry point and the caller cannot apply half of it.
    /// </summary>
    public sealed class OverlayStyle
    {
        public OverlayStyle(ThemePalette palette, int fontSize, OverlayLayout layout)
        {
            Palette = palette;
            FontSize = fontSize;
            Layout = layout;
        }

        public ThemePalette Palette { get; }

        public int FontSize { get; }

        public OverlayLayout Layout { get; }
    }
}
