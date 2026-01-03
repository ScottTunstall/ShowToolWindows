using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace ShowToolWindows.UI
{
    /// <summary>
    /// Provides helper methods for opening Visual Studio tool windows.
    /// </summary>
    internal class ToolWindowHelper
    {
        private readonly DTE _dte;
        private readonly IVsUIShell _uiShell;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowHelper"/> class.
        /// </summary>
        /// <param name="dte">The Visual Studio automation object.</param>
        /// <param name="vsUiShell">The Visual Studio UI shell service.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="dte"/> or <paramref name="vsUiShell"/> is null.</exception>
        public ToolWindowHelper(DTE dte, IVsUIShell vsUiShell)
        {
            _dte = dte ?? throw new ArgumentNullException(nameof(dte));
            _uiShell = vsUiShell ?? throw new ArgumentNullException(nameof(vsUiShell));
        }


        /// <summary>
        /// Attempts to open a tool window by its object kind.
        /// </summary>
        /// <param name="objectKind">The tool window object kind GUID.</param>
        /// <returns><c>true</c> if the window is opened or already available; otherwise, <c>false</c>.</returns>
        public bool TryOpenToolWindowByObjectKind(string objectKind)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var window = _dte.Windows.Item(objectKind);
                if (window != null)
                {
                    window.Visible = true;
                    window.Activate();
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Window '{objectKind}' not in collection, will try to create it: {ex.Message}");
            }

            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                try
                {
                    Guid toolWindowGuid = new Guid(objectKind);
                    int hr = _uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref toolWindowGuid, out IVsWindowFrame frame);
                    if (ErrorHandler.Succeeded(hr) && frame != null)
                    {
                        frame.Show();
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not open window ObjectKind '{objectKind}' via IVsUIShell: {ex.Message}");
                }

                return false;
            });

            return false;
        }


    }
}
