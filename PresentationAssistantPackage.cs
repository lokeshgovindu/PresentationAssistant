using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using PresentationAssistant.State;
using PresentationAssistant.Theming;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace PresentationAssistant
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [Guid(PresentationAssistantPackage.PackageGuidString)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.NoSolution_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideOptionPage(typeof(PresentationAssistantOptionsDialog),
        PresentationAssistantOptionsDialog.Category, PresentationAssistantOptionsDialog.SubCategory,
        1000, 1001, true)]
    [ProvideProfile(typeof(PresentationAssistantOptionsDialog),
        PresentationAssistantOptionsDialog.Category, PresentationAssistantOptionsDialog.SubCategory,
        1000, 1001, true)]
    public sealed class PresentationAssistantPackage : AsyncPackage
    {
        /// <summary>
        /// PresentationAssistantPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "21f2f4b1-873b-4456-ba7b-2101d1a686c9";

        public const string ApplicationName      = Product.Name;
        public const string ApplicationNameShort = Product.ShortName;

        #region Package Members

        private static readonly object              _windowLock = new object();

        private static DTE2                         _dte;
        private static CommandEvents                _commandEvents;
        private static PresentationAssistantWindow  _window = null;
        private static ShortcutDisplayStatistics    _statistics;
        private static Settings                     _settings;
        private static CommandExclusions            _exclusions = CommandExclusions.Empty;

        private static OutputWindowPane             _outputWindowPane = null;
        private static readonly string              OutputWindowName  = ApplicationName;

        public static Settings Settings => _settings;

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var shellService = await this.GetServiceAsync(typeof(SVsShell)) as IVsShell;
            if (shellService != null)
            {
                InitializeServices();
            }

            if (_dte == null)
            {
                Debug.WriteLine($"[{ApplicationName}] No DTE, giving up on initialization.");
                return;
            }

            _settings   = Settings.Load();
            _exclusions = new CommandExclusions(_settings.ExcludedCommands);
            _statistics = new ShortcutDisplayStatistics(_settings.MultiplierTimeoutInMS);

            _commandEvents = _dte.Events.CommandEvents;
            _commandEvents.BeforeExecute += CommandEvents_BeforeExecute;

            Settings.SettingsUpdated  += OnSettingsUpdated;
            ThemeManager.ThemeChanged += OnShellThemeChanged;

            OutputLine("{0} initialized. Theme={1}, WindowTimeout={2}ms, MultiplierTimeout={3}ms, ShortcutsOnly={4}",
                ApplicationName, _settings.Theme, _settings.WindowTimeoutInMS,
                _settings.MultiplierTimeoutInMS, _settings.ShortcutsOnly);
        }

        private void InitializeServices()
        {
            _dte = this.GetService<SDTE, SDTE>() as DTE2;

            Debug.Assert(_dte != null, "dte != null");
            if (_dte == null)
            {
                Debug.WriteLine($"[{ApplicationName}] Cannot get a DTE service.");
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                Settings.SettingsUpdated  -= OnSettingsUpdated;
                ThemeManager.ThemeChanged -= OnShellThemeChanged;

                // Dispose can run off the main thread during shutdown, and the DTE event
                // object may only be touched from it - so skip the unsubscribe rather
                // than marshalling (which risks deadlocking a shutting-down shell).
                if (_commandEvents != null && ThreadHelper.CheckAccess())
                {
#pragma warning disable VSTHRD010 // CheckAccess above already established main-thread affinity.
                    _commandEvents.BeforeExecute -= CommandEvents_BeforeExecute;
#pragma warning restore VSTHRD010
                }
            }

            base.Dispose(disposing);
        }

        #endregion

        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_settings == null || !_settings.Enabled) return;

            var shortcut = GetShortcut(Guid, ID);
            if (shortcut == null) return;
            if (_settings.ShortcutsOnly && !shortcut.HasShortcuts) return;

            // Count the invocation only once the command has passed every filter,
            // otherwise a suppressed command in the middle of a run would reset the
            // multiplier.
            _statistics.OnAction(shortcut.ActionId);
            shortcut.Multiplier = _statistics.Multiplier;

            OutputLine("Guid: {0}, ID: {1}, {2}", Guid, ID, shortcut);

            try
            {
                ShowShortcut(shortcut);
            }
            catch (Exception ex)
            {
                // Visual Studio swallows exceptions thrown out of DTE event handlers, so
                // report them here rather than letting the overlay fail invisibly.
                Debug.WriteLine($"[{ApplicationName}] Failed to show the overlay: {ex}");
                OutputLine("Failed to show the overlay: {0}", ex);
            }
        }

        private static void ShowShortcut(ShortcutDetails shortcut)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            lock (_windowLock)
            {
                if (_window == null)
                {
                    _window = new PresentationAssistantWindow(_settings.WindowTimeoutInMS);

                    // If the window ever does close - the shell tearing down its owner,
                    // say - drop it so the next command builds a fresh one instead of
                    // calling Show() on a closed window forever.
                    _window.Closed += OnWindowClosed;
                }

                // Re-resolving per show is what makes an edit to themes.json visible on the
                // next keystroke. The catalog only re-reads the file when its timestamp
                // changed, so this is a dictionary lookup in the common case.
                _window.ShowShortcut(shortcut, CurrentStyle());

                _dte.StatusBar.Text = shortcut.HasShortcuts
                    ? $"{ApplicationNameShort}: {shortcut.ActionId} via {shortcut.ShortcutsStr}"
                    : $"{ApplicationNameShort}: {shortcut.ActionId}";
            }
        }

        private ShortcutDetails GetShortcut(string Guid, int ID)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Command cmd;
            try { cmd = _dte.Commands.Item(Guid, ID); } catch (Exception) { return null; }
            if (cmd == null) return null;

            // Filtering must use the canonical name: LocalizedName is translated, and the
            // blocklists are written in canonical form, so matching on the localized name
            // silently stops working on a non-English IDE.
            string actionId = GetCanonicalName(cmd);
            if (string.IsNullOrEmpty(actionId)) return null;
            if (ActionIdBlocklist.IsBlocked(actionId)) return null;
            if (_exclusions.IsExcluded(actionId)) return null;

            string[] shortcuts = null;
            if (cmd.Bindings is object[] bindings && bindings.Any())
            {
                shortcuts = GetBindings(bindings);
                if (shortcuts.Length == 0) shortcuts = null;
            }

            return new ShortcutDetails
            {
                ActionId = actionId,

                // The description is what the user reads, so prefer the localized name.
                Description = GetCommandDescription(GetDisplayName(cmd, actionId)),
                Shortcuts   = shortcuts,
                Multiplier  = 1
            };
        }

        /// <summary>
        /// Opens themes.json for editing, creating it and refreshing the generated
        /// reference listing first. Driven by the "..." button on the options page.
        /// </summary>
        public static void OpenThemesFile()
        {
            ThemeCatalog.EnsureAuthoringFiles();

            // Deliberately NOT the IDE's own editor. This is invoked from the options
            // page, which is a modal dialog: a document opened into the IDE appears behind
            // it, so the button looks like it did nothing. Notepad is always present and
            // always comes to the front, and plain JSON needs nothing more.
            try
            {
                System.Diagnostics.Process.Start("notepad.exe", "\"" + AppPaths.ThemesFile + "\"");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{ApplicationName}] Failed to open themes.json in Notepad: {ex}");

                try
                {
                    System.Diagnostics.Process.Start(
                        new ProcessStartInfo(AppPaths.ThemesFile) { UseShellExecute = true });
                }
                catch (Exception fallback)
                {
                    Debug.WriteLine($"[{ApplicationName}] Failed to open themes.json: {fallback}");
                }
            }
        }

        private static OverlayStyle CurrentStyle()
        {
            return new OverlayStyle(
                ThemeManager.Resolve(_settings.Theme),
                _settings.FontSize,
                _settings.Layout);
        }

        private static void OnWindowClosed(object sender, EventArgs e)
        {
            if (ReferenceEquals(sender, _window)) _window = null;
        }

        private void OnSettingsUpdated(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _settings = Settings.Load();
            _exclusions = new CommandExclusions(_settings.ExcludedCommands);
            _statistics.SetMultiplierTimeout(_settings.MultiplierTimeoutInMS);

            // The user may have edited themes.json while the dialog was open.
            ThemeCatalog.Invalidate();

            // The overlay is created lazily, on the first announced command.
            if (_window != null)
            {
                _window.SetWindowTimeout(_settings.WindowTimeoutInMS);
                _window.ApplyStyle(CurrentStyle());
            }
        }

        private void OnShellThemeChanged(object sender, EventArgs e)
        {
            // Only Auto and VisualStudio derive from the shell, but re-resolving a fixed
            // palette is a dictionary lookup, so don't bother special-casing.
            if (_window != null && _settings != null)
            {
                _window.ApplyStyle(CurrentStyle());
            }
        }

        private static string[] GetBindings(IEnumerable<object> bindings)
        {
            // Bindings arrive as "Scope::Key", e.g. "Text Editor::Ctrl+Down Arrow".
            var result = bindings.Select(binding => binding.ToString().IndexOf("::") >= 0
                ? binding.ToString().Substring(binding.ToString().IndexOf("::") + 2)
                : binding.ToString()).Distinct();

            return result.ToArray();
        }

        /// <summary>The untranslated command name, e.g. <c>Edit.ScrollLineDown</c>.</summary>
        private static string GetCanonicalName(Command vsCommand)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return string.IsNullOrWhiteSpace(vsCommand.Name) ? vsCommand.LocalizedName : vsCommand.Name;
        }

        /// <summary>The translated command name, falling back to the canonical one.</summary>
        private static string GetDisplayName(Command vsCommand, string canonical)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return string.IsNullOrWhiteSpace(vsCommand.LocalizedName) ? canonical : vsCommand.LocalizedName;
        }

        private static string GetCommandDescription(string actionId)
        {
            string commandName = actionId.Substring(actionId.LastIndexOf('.') + 1);

            // Define a regular expression pattern to split the string at capital letters
            string pattern = @"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])";

            // Split the input into words using the regular expression pattern
            string[] words = Regex.Split(commandName, pattern);

            return String.Join(" ", words);
        }

        /// <summary>
        /// Writes a line to the PresentationAssistant pane in the Output window when the
        /// Diagnostics setting is on. Available in Release too: without it a user reporting
        /// a problem has nothing to send back. The pane is created on first use, so it
        /// costs nothing while diagnostics are off.
        /// </summary>
        public static void OutputLine(string messageFormat, params object[] arguments)
        {
            if (_settings == null || !_settings.Diagnostics) return;
            if (!ThreadHelper.CheckAccess()) return;

            try
            {
#pragma warning disable VSTHRD010 // CheckAccess above already established main-thread affinity.
                var pane = GetOutputWindowPane();
                var message = arguments == null || arguments.Length == 0
                    ? messageFormat
                    : string.Format(messageFormat, arguments);

                pane?.OutputString(message + Environment.NewLine);
#pragma warning restore VSTHRD010
            }
            catch (Exception ex)
            {
                // Diagnostics must never be the thing that breaks the extension.
                Debug.WriteLine($"[{ApplicationName}] Failed to write diagnostics: {ex}");
            }
        }

        private static OutputWindowPane GetOutputWindowPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_outputWindowPane == null && _dte != null)
            {
                var window = _dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
                var outputWindow = (OutputWindow)window.Object;

                // The analyzer cannot see the assertion above from inside the lambda.
#pragma warning disable VSTHRD010
                _outputWindowPane = outputWindow.OutputWindowPanes
                    .Cast<OutputWindowPane>()
                    .FirstOrDefault(p => p.Name == OutputWindowName)
                    ?? outputWindow.OutputWindowPanes.Add(OutputWindowName);
#pragma warning restore VSTHRD010
            }

            return _outputWindowPane;
        }
    }
}
