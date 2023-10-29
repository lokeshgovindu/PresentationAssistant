using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
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
    public sealed class PresentationAssistantPackage : AsyncPackage
    {
        /// <summary>
        /// PresentationAssistantPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "21f2f4b1-873b-4456-ba7b-2101d1a686c9";

        #region Package Members

        private static DTE2                         _dte;
        private static IAsyncServiceProvider        _serviceProvider;
        private static CommandEvents                _commandEvents;
        private static PresentationAssistantWindow  _window           = null;

#if DEBUG
        private static OutputWindowPane             _outputWindowPane = null;
        private static string                       OutputWindowName  = "PresentationAssistant";
#endif

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

#if DEBUG
            CreateOutputWindowPane();
#endif

            _commandEvents = _dte.Events.CommandEvents;
            _commandEvents.BeforeExecute += CommandEvents_BeforeExecute;
        }

        private void InitializeServices()
        {
            _dte = this.GetService<SDTE, SDTE>() as DTE2;

            Debug.Assert(_dte != null, "dte != null");
            if (_dte != null)
            {
                PresentationAssistantPackage._serviceProvider = this;
            }
            else
            {
                Debug.WriteLine("[PresentationAssistant] Cannot get a DTE service.");
            }
        }

#endregion

        private void CommandEvents_BeforeExecute(string Guid, int ID, object CustomIn, object CustomOut, ref bool CancelDefault)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Command cmd = null;
            try { cmd = _dte.Commands.Item(Guid, ID); } catch (Exception) { return; }
            if (cmd == null) return;

            var bindings = cmd.Bindings as object[];
            if (bindings == null && !bindings.Any()) return;

            var shortcuts = GetBindings(bindings);
            if (!shortcuts.Any()) return;

            string actionId = GetCommandName(cmd);
            bool isBlocked = ActionIdBlocklist.IsBlocked(actionId);

#if DEBUG
            _outputWindowPane.OutputString(
                String.Format("Command: {0}, IsBlocked: {1}, Shortcuts: {2}\n",
                    actionId, isBlocked, string.Join(" || ", shortcuts)));
            if (_window != null) {
                _outputWindowPane.OutputString(
                    String.Format("Window.ActionId = {0}, Window.Terminated = {1}\n",
                        _window.ActionId, _window.Terminated));
            }
#endif
            if (isBlocked) return;
            if (_window != null && _window.ActionId == actionId && !_window.Terminated) return;

            lock (typeof(PresentationAssistantPackage)) {
#if DEBUG
                _outputWindowPane.OutputString("Inside lock (typeof(PresentationAssistantPackage))\n");
#endif
                if (_window != null) {
                    _window.Close();
                    _window = null;
                }

                _window = new PresentationAssistantWindow(actionId);
                var contentBlock = _window.CommandText.Children;
                contentBlock.Clear();
                Run actionIdRun = new Run(actionId);
                actionIdRun.FontWeight = FontWeights.Bold;
                contentBlock.Add(new TextBlock(actionIdRun));
                contentBlock.Add(new TextBlock(new Run(" via ")));
                var space = Convert.ToChar(160);
                for (int i = 0; i < shortcuts.Length; ++i)
                {
                    if (i > 0) {
                        contentBlock.Add(new TextBlock(new Run(space + "or" + space)) { Opacity = 0.5 });
                    }
                    contentBlock.Add(new TextBlock(new Run(shortcuts[i])));
                }
                _window.Show();
            }
        }

        private static string[] GetBindings(IEnumerable<object> bindings)
        {
            var result = bindings.Select(binding => binding.ToString().IndexOf("::") >= 0
                ? binding.ToString().Substring(binding.ToString().IndexOf("::") + 2)
                : binding.ToString()).Distinct();

            return result.ToArray();
        }

        private static string GetCommandName(Command vsCommand)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return string.IsNullOrWhiteSpace(vsCommand.LocalizedName) ? vsCommand.Name : vsCommand.LocalizedName;
        }

#if DEBUG
        public static OutputWindowPane GetOutputWindowPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_outputWindowPane == null)
            {
                var outputWindow = (OutputWindow)GetOutputWindow().Object;
                _outputWindowPane = outputWindow.OutputWindowPanes.Add(OutputWindowName);
            }
            return _outputWindowPane;
        }

        public static void CreateOutputWindowPane()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (_outputWindowPane == null)
            {
                var outputWindow = (OutputWindow)GetOutputWindow().Object;
                _outputWindowPane = outputWindow.OutputWindowPanes.Add(OutputWindowName);
                _outputWindowPane.OutputString("Output window created\n");
            }
        }

        public static EnvDTE.Window GetOutputWindow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            return _dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);
        }
#endif
    }
}
