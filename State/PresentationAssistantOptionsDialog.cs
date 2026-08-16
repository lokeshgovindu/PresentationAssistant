using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Windows;

namespace PresentationAssistant.State
{
    /// <summary>
    /// Tools &gt; Options &gt; PresentationAssistant &gt; General.
    /// </summary>
    /// <remarks>
    /// A hand-built WPF page rather than the default property grid, because the settings
    /// that matter most are visual: a grid cannot show what a theme, font size or layout
    /// will actually look like, and it turns the exclusion list into one long line. The
    /// page previews the overlay live as the settings change.
    /// <para>
    /// Backed by <see cref="Settings"/> rather than the shell's own property storage, so
    /// the JSON file stays the single source of truth. The type name is unchanged from the
    /// grid-based version it replaced, which keeps the registered page GUID stable.
    /// </para>
    /// </remarks>
    [Serializable]
    internal class PresentationAssistantOptionsDialog : UIElementDialogPage
    {
        public const string Category    = "PresentationAssistant";
        public const string SubCategory = "General";

        private OptionsViewModel _viewModel;
        private OptionsPageControl _control;

        protected override UIElement Child
        {
            get
            {
                if (_control == null)
                {
                    _control = new OptionsPageControl { DataContext = ViewModel };
                }

                return _control;
            }
        }

        private OptionsViewModel ViewModel => _viewModel ?? (_viewModel = new OptionsViewModel());

        /// <summary>Re-read on every activation, so hand edits to the file are picked up.</summary>
        protected override void OnActivate(CancelEventArgs e)
        {
            base.OnActivate(e);
            ViewModel.Load(Settings.Load());
        }

        public override void LoadSettingsFromStorage()
        {
            ViewModel.Load(Settings.Load());
        }

        public override void SaveSettingsToStorage()
        {
            ViewModel.ToSettings().Save();
        }

        /// <summary>OK or Apply. Saving raises SettingsUpdated, which re-styles a live overlay.</summary>
        protected override void OnApply(PageApplyEventArgs e)
        {
            if (e.ApplyBehavior == ApplyKind.Apply)
            {
                // Colours first: saving settings raises SettingsUpdated, which re-resolves
                // the theme, so the override has to be on disk by then.
                ViewModel.SaveColoursIfEdited();
                ViewModel.ToSettings().Save();
            }

            base.OnApply(e);
        }
    }
}
