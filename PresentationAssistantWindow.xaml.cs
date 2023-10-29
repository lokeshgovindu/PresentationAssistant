using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Application = System.Windows.Application;

namespace PresentationAssistant
{
    /// <summary>
    /// Interaction logic for PresentationAssistantWindow.xaml
    /// </summary>
    public partial class PresentationAssistantWindow : Window
    {
        private Window ParentWindow;
        public  string ActionId     { get; private set; }
        public  bool   Terminated   { get; private set; }

        private readonly int LifeTimeInMS = 2000;

        public PresentationAssistantWindow(string actionId)
        {
            InitializeComponent();

            // To avoid showing in Alt+Tab
            this.Owner = Application.Current.MainWindow;
            this.Terminated = false;

            //this.ParentWindow = Application.Current.MainWindow;
            FindParentWindow();

            this.Loaded += PresentationAssistantWindow_Loaded;

            _ = CloseAfterSomeTimeAsync();
            ActionId = actionId;
        }

        private void PresentationAssistantWindow_Loaded(object sender, RoutedEventArgs e)
        {
            PlaceWindow();
        }

        public async Task CloseAfterSomeTimeAsync()
        {
            await Task.Delay(LifeTimeInMS);
            Close();
            Terminated = true;
        }

        public void PlaceWindow()
        {
            if (this.ParentWindow != null)
            {
                double parentLeft = this.ParentWindow.Left;

                if (this.ParentWindow.WindowState == WindowState.Maximized)
                {
                    var windowInterop = new WindowInteropHelper(this.ParentWindow);
                    var screen = Screen.FromHandle(windowInterop.Handle);

                    parentLeft = screen.WorkingArea.Left;
                }

                // StatusBar height - 23
                this.Top = (this.ParentWindow.ActualHeight - this.ActualHeight - 31);
                this.Left = parentLeft + ((this.ParentWindow.ActualWidth - this.ActualWidth) / 2);
            }
        }

        private void FindParentWindow()
        {
            this.ParentWindow = Application.Current.Windows
                .Cast<object>()
                .Where(w => w is Window && !Equals(w, this))
                .Cast<Window>()
                .FirstOrDefault(window => window.IsActive) ?? Application.Current.MainWindow;
        }
    }
}
