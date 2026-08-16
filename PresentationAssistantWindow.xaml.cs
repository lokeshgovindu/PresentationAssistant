using PresentationAssistant.Theming;
using System;
using System.Runtime.InteropServices;
using System.Windows.Media;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using Application = System.Windows.Application;
using Screen = System.Windows.Forms.Screen;

namespace PresentationAssistant
{
    /// <summary>
    /// The borderless overlay shown above the status bar. A single instance is reused
    /// for every command: <see cref="SetShortcut"/> swaps the bound data and
    /// <see cref="ReShow"/> restarts the hide timer.
    /// </summary>
    public partial class PresentationAssistantWindow : Window
    {
        private Window _parentWindow;
        private CancellationTokenSource _cancellationTokenSource;

        /// <summary>
        /// Incremented on every show. A pending hide whose generation is stale has been
        /// superseded by a newer command and must not touch the window.
        /// </summary>
        private int _showGeneration;

        /// <summary>
        /// Where the overlay waits while idle: far enough left to be off any monitor, but
        /// still a mapped, composing window. See <see cref="ShowShortcut"/>.
        /// </summary>
        private const double ParkedLeft = -32000;

        /// <summary>False until the first <see cref="Show"/>.</summary>
        private bool _hasBeenShown;

        /// <summary>True while the overlay is parked off-screen rather than on display.</summary>
        private bool _parked;

        private bool _revealPending;
        private int _framesWaited;

        public PresentationAssistantWindow(int windowTimeoutInMS)
        {
            InitializeComponent();

            // Owned windows stay out of Alt+Tab.
            Owner = Application.Current.MainWindow;
            Terminated = false;

            FindParentWindow();

            Loaded      += PresentationAssistantWindow_Loaded;
            SizeChanged += PresentationAssistantWindow_SizeChanged;

            SetWindowTimeout(windowTimeoutInMS);
            _cancellationTokenSource = new CancellationTokenSource();

            // The hide timer is armed by ShowShortcut, not here: nothing is on screen yet.
        }

        /// <summary>True once the hide timer has elapsed and the overlay was hidden.</summary>
        public bool Terminated { get; private set; }

        public ShortcutDetails Shortcut { get; private set; }

        public TimeSpan WindowTimeout { get; private set; } = TimeSpan.FromMilliseconds(5000);

        public void SetShortcut(ShortcutDetails shortcut)
        {
            Shortcut = shortcut;
            DataContext = shortcut;
        }

        public void SetWindowTimeout(int windowTimeoutInMS)
        {
            WindowTimeout = TimeSpan.FromMilliseconds(windowTimeoutInMS);
        }

        /// <summary>
        /// Swaps in the colors of <paramref name="palette"/>. The XAML binds these keys
        /// with DynamicResource, so replacing the entries repaints the live window.
        /// </summary>
        public void ApplyTheme(ThemePalette palette)
        {
            if (palette == null) return;

            Resources[ThemePalette.BackgroundKey]          = palette.Background;
            Resources[ThemePalette.ForegroundKey]          = palette.Foreground;
            Resources[ThemePalette.SecondaryForegroundKey] = palette.SecondaryForeground;
            Resources[ThemePalette.BorderKey]              = palette.Border;

            Opacity = palette.Opacity;
        }

        /// <summary>
        /// Announces <paramref name="shortcut"/>: applies the palette, swaps the content,
        /// brings the overlay into view and restarts the display timer.
        /// </summary>
        /// <remarks>
        /// <para>
        /// The overlay is reused between commands rather than recreated, which costs about
        /// 3ms per command instead of 17ms. The catch is that WPF does not compose a hidden
        /// window: it keeps the last frame it drew while visible. Hiding it and showing it
        /// again therefore presented the *previous* command for a frame or two before the
        /// new content was composed - the flash this method exists to avoid.
        /// </para>
        /// <para>
        /// So the window is never hidden. When idle it is parked off-screen, where it stays
        /// mapped and WPF keeps composing it. A new command updates the content while it is
        /// still parked, waits for that content to actually be composed, and only then
        /// moves it into view - so the first frame the user sees is always correct.
        /// </para>
        /// </remarks>
        public void ShowShortcut(ShortcutDetails shortcut, ThemePalette palette)
        {
            ApplyTheme(palette);
            SetShortcut(shortcut);

            if (!_hasBeenShown)
            {
                // Map it off-screen; the reveal below brings it in once it has a frame.
                Park();
                Show();
                _hasBeenShown = true;
            }

            // Flush bindings and layout now so the window is the right size, and so the
            // frame composed while parked is the one we want to reveal.
            UpdateLayout();

            RestartHideTimer();

            if (_parked)
            {
                RevealWhenComposed();
            }
            else
            {
                // Already on screen: this is a further command in the same run, so updating
                // in place is what the user expects. Re-place it in case the text width
                // changed or the IDE moved.
                PlaceWindow();
            }
        }

        /// <summary>
        /// Makes the overlay transparent to the mouse. It sits over the editor for several
        /// seconds at a time, and being a normal window it would otherwise swallow every
        /// click that landed on it - a dead strip above the status bar. WS_EX_TRANSPARENT
        /// takes it out of hit testing so clicks reach the IDE underneath, and
        /// WS_EX_NOACTIVATE keeps it from ever taking activation.
        /// </summary>
        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);

            var handle = new WindowInteropHelper(this).Handle;
            if (handle == IntPtr.Zero) return;

            var style = GetExtendedStyle(handle).ToInt64();
            SetExtendedStyle(handle, new IntPtr(style | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE));
        }

        /// <summary>Moves the overlay out of sight without unmapping it.</summary>
        private void Park()
        {
            CancelPendingReveal();

            _parked = true;
            Left = ParkedLeft;
        }

        /// <summary>
        /// Moves the overlay into view once the content set while parked has actually been
        /// composed. Rendering fires immediately before each frame, so the frame containing
        /// the new content is the one after the first tick.
        /// </summary>
        private void RevealWhenComposed()
        {
            if (_revealPending) return;

            _revealPending = true;
            _framesWaited = 0;
            CompositionTarget.Rendering += OnRenderingBeforeReveal;
        }

        private void OnRenderingBeforeReveal(object sender, EventArgs e)
        {
            if (++_framesWaited < 2) return;

            CancelPendingReveal();
            PlaceWindow();
        }

        private void CancelPendingReveal()
        {
            if (!_revealPending) return;

            CompositionTarget.Rendering -= OnRenderingBeforeReveal;
            _revealPending = false;
        }

        private void RestartHideTimer()
        {
            if (!Terminated)
            {
                _cancellationTokenSource.Cancel();
            }

            _cancellationTokenSource.Dispose();
            _cancellationTokenSource = new CancellationTokenSource();

            Terminated = false;
            _ = HideAfterTimeoutAsync();
        }

        /// <summary>
        /// Centres the overlay horizontally, sits it just above the status bar, and brings
        /// it out of the parked position.
        /// </summary>
        public void PlaceWindow()
        {
            // Re-resolve rather than give up: a parked window that never gets placed would
            // leave the overlay permanently invisible.
            if (_parentWindow == null) FindParentWindow();

            if (_parentWindow == null)
            {
                PlaceOnPrimaryScreen();
                return;
            }

            var parentTop  = _parentWindow.Top;
            var parentLeft = _parentWindow.Left;

            if (_parentWindow.WindowState == WindowState.Maximized)
            {
                // A maximized window reports its restored bounds, so ask the screen
                // instead. Screen bounds are device pixels while Top/Left are device
                // independent units, so convert - otherwise the overlay is misplaced on a
                // scaled display.
                var screen = Screen.FromHandle(new WindowInteropHelper(_parentWindow).Handle);
                var corner = new Point(screen.WorkingArea.Left, screen.WorkingArea.Top);

                var source = PresentationSource.FromVisual(_parentWindow);
                if (source?.CompositionTarget != null)
                {
                    corner = source.CompositionTarget.TransformFromDevice.Transform(corner);
                }

                parentTop  = corner.Y;
                parentLeft = corner.X;
            }

            // 31 leaves room for the status bar.
            Top  = parentTop + (_parentWindow.ActualHeight - ActualHeight - 31);
            Left = parentLeft + (_parentWindow.ActualWidth - ActualWidth) / 2;

            _parked = false;
        }

        /// <summary>Fallback placement when there is no shell window to hang off.</summary>
        private void PlaceOnPrimaryScreen()
        {
            var area = Screen.PrimaryScreen.WorkingArea;
            var corner = new Point(area.Left, area.Bottom);

            var source = PresentationSource.FromVisual(this);
            if (source?.CompositionTarget != null)
            {
                corner = source.CompositionTarget.TransformFromDevice.Transform(corner);
            }

            Top  = corner.Y - ActualHeight - 31;
            Left = corner.X + (area.Width - ActualWidth) / 2;

            _parked = false;
        }

        private async Task HideAfterTimeoutAsync()
        {
            var generation = ++_showGeneration;
            var token = _cancellationTokenSource.Token;

            try
            {
                await Task.Delay(WindowTimeout, token);
            }
            catch (OperationCanceledException)
            {
                return;
            }

            // A newer command was announced while we were waiting.
            if (generation != _showGeneration) return;

            // Park rather than Hide: a hidden window stops being composed, and the stale
            // frame is what caused the overlay to flash the previous command. See ShowShortcut.
            Park();
            Terminated = true;
        }

        private void PresentationAssistantWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // Never place while parked - that would defeat the wait for a composed frame.
            if (!_parked) PlaceWindow();
        }

        private void PresentationAssistantWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Fires from the very first Show(), which happens while parked on purpose.
            if (!_parked) PlaceWindow();
        }

        private void FindParentWindow() => _parentWindow = Application.Current?.MainWindow;

        #region Click-through interop

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_NOACTIVATE = 0x08000000;

        // Visual Studio is 64-bit, but pick the right entry point rather than assume it.
        private static IntPtr GetExtendedStyle(IntPtr handle)
        {
            return IntPtr.Size == 8
                ? GetWindowLongPtr64(handle, GWL_EXSTYLE)
                : new IntPtr(GetWindowLong32(handle, GWL_EXSTYLE));
        }

        private static void SetExtendedStyle(IntPtr handle, IntPtr style)
        {
            if (IntPtr.Size == 8)
            {
                SetWindowLongPtr64(handle, GWL_EXSTYLE, style);
            }
            else
            {
                SetWindowLong32(handle, GWL_EXSTYLE, style.ToInt32());
            }
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

        #endregion
    }
}
