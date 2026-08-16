using PresentationAssistant.Theming;
using System.Windows;
using System.Windows.Controls;
using Forms = System.Windows.Forms;
using Media = System.Windows.Media;

namespace PresentationAssistant.State
{
    /// <summary>
    /// The options UI. Deliberately thin: state lives on <see cref="OptionsViewModel"/>,
    /// which the hosting page supplies as the DataContext. The only logic here is what
    /// cannot be expressed in XAML - opening dialogs.
    /// </summary>
    public partial class OptionsPageControl : UserControl
    {
        public OptionsPageControl()
        {
            InitializeComponent();
        }

        private OptionsViewModel ViewModel => DataContext as OptionsViewModel;

        private void EditThemesButton_Click(object sender, RoutedEventArgs e)
        {
            PresentationAssistantPackage.OpenThemesFile();
        }

        private void ResetColoursButton_Click(object sender, RoutedEventArgs e)
        {
            ViewModel?.ResetColours();
        }

        /// <summary>
        /// Picks a colour with the system dialog. WPF has no colour picker of its own, and
        /// the Windows Forms one is already available, native, and familiar.
        /// </summary>
        private void Swatch_Click(object sender, RoutedEventArgs e)
        {
            var model = ViewModel;
            var slot = (sender as FrameworkElement)?.Tag as string;
            if (model == null || slot == null) return;

            var current = CurrentHex(model, slot);

            using (var dialog = new Forms.ColorDialog { FullOpen = true, AnyColor = true })
            {
                if (ThemePalette.TryParse(current, out var parsed))
                {
                    dialog.Color = System.Drawing.Color.FromArgb(parsed.R, parsed.G, parsed.B);
                }

                if (dialog.ShowDialog() != Forms.DialogResult.OK) return;

                var picked = ThemePalette.ToHex(
                    Media.Color.FromRgb(dialog.Color.R, dialog.Color.G, dialog.Color.B));

                Apply(model, slot, picked);
            }
        }

        private static string CurrentHex(OptionsViewModel model, string slot)
        {
            switch (slot)
            {
                case "Background": return model.BackgroundHex;
                case "Foreground": return model.ForegroundHex;
                case "Secondary":  return model.SecondaryHex;
                case "Border":     return model.BorderHex;
                default:           return null;
            }
        }

        private static void Apply(OptionsViewModel model, string slot, string hex)
        {
            switch (slot)
            {
                case "Background": model.BackgroundHex = hex; break;
                case "Foreground": model.ForegroundHex = hex; break;
                case "Secondary":  model.SecondaryHex = hex; break;
                case "Border":     model.BorderHex = hex; break;
            }
        }
    }
}
