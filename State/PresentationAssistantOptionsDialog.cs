using Microsoft.VisualStudio.Shell;
using PresentationAssistant.Theming;
using System;
using System.ComponentModel;

namespace PresentationAssistant.State
{
    /// <summary>
    /// Tools &gt; Options &gt; PresentationAssistant &gt; General. Backed by
    /// <see cref="Settings"/> rather than by the shell's own property storage, so the
    /// JSON file stays the single source of truth.
    /// </summary>
    [Serializable]
    internal class PresentationAssistantOptionsDialog : DialogPage
    {
        public const string Category    = "PresentationAssistant";
        public const string SubCategory = "General";

        [System.ComponentModel.Category("General")]
        [DisplayName("Window Timeout (MS)")]
        [Description("Window timeout in milliseconds, hide window after timespan")]
        public int WindowTimeout { get; set; }

        [System.ComponentModel.Category("General")]
        [DisplayName("Multiplier Timeout (MS)")]
        [Description("Multiplier timeout in milliseconds, the maximum gap between two key presses that still count as a repeat")]
        public int MultiplierTimeout { get; set; }

        [System.ComponentModel.Category("General")]
        [DisplayName("Shortcuts Only")]
        [Description("Show commands having shortcuts only")]
        public bool ShortcutsOnly { get; set; }

        [System.ComponentModel.Category("General")]
        [DisplayName("Excluded Commands")]
        [Description("Commands never to announce, separated by semicolons. Matched against the command name shown in the status bar; end a pattern with * to match a prefix, e.g. \"Edit.Line*; View.Output\". A set of noisy IDE commands is always excluded on top of this.")]
        public string ExcludedCommands { get; set; }

        [System.ComponentModel.Category("Appearance")]
        [DisplayName("Theme")]
        [Description("Overlay colors. Auto follows the Visual Studio theme, VisualStudio takes the shell colors as they are, and the rest are palettes - including any you define in themes.json.")]
        [TypeConverter(typeof(ThemeNameConverter))]
        public string Theme { get; set; }

        [System.ComponentModel.Category("Appearance")]
        [DisplayName("Themes File")]
        [Description("Where custom themes live. Press the \"...\" button to open it, together with themes.reference.json listing every built-in theme's colours to copy from. Changes apply on the next keystroke.")]
        [Editor(typeof(ThemesFileEditor), typeof(System.Drawing.Design.UITypeEditor))]
        public string ThemesFile
        {
            get { return AppPaths.ThemesFile; }

            // The grid only offers the "..." button on a settable property; the path is
            // derived, so anything written here is discarded.
            set { }
        }

        public override void LoadSettingsFromStorage()
        {
            var settings = Settings.Load();

            WindowTimeout     = settings.WindowTimeoutInMS;
            MultiplierTimeout = settings.MultiplierTimeoutInMS;
            ShortcutsOnly     = settings.ShortcutsOnly;
            Theme             = settings.Theme;

            // Stored as a JSON array so the file stays readable; shown as one editable
            // line, because the property grid's collection editor is far more awkward.
            ExcludedCommands = CommandExclusions.Join(settings.ExcludedCommands);
        }

        public override void SaveSettingsToStorage()
        {
            new Settings
            {
                WindowTimeoutInMS     = WindowTimeout,
                MultiplierTimeoutInMS = MultiplierTimeout,
                ShortcutsOnly         = ShortcutsOnly,
                Theme                 = Theme,
                ExcludedCommands      = CommandExclusions.Split(ExcludedCommands)
            }.Save();
        }
    }
}
