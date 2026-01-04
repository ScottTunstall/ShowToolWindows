using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using System;
using System.Collections.Generic;

namespace ShowToolWindows.UI
{
    /// <summary>
    /// Provides helper methods for opening and managing Visual Studio tool windows.
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
        /// Sets the visibility state of a single tool window.
        /// </summary>
        /// <param name="entry">The tool window entry to modify.</param>
        /// <param name="isVisible">Whether the tool window should be visible.</param>
        public void SetToolWindowVisibility(ToolWindowEntry entry, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            entry.SetVisibility(isVisible);
            entry.Synchronize();
        }

        /// <summary>
        /// Sets the visibility state of multiple tool windows.
        /// </summary>
        /// <param name="toolWindows">The collection of tool window entries to modify.</param>
        /// <param name="isVisible">Whether the tool windows should be visible.</param>
        public void SetToolWindowsVisibility(IEnumerable<ToolWindowEntry> toolWindows, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (ToolWindowEntry entry in toolWindows)
            {
                entry.SetVisibility(isVisible);
                entry.Synchronize();
            }
        }

        /// <summary>
        /// Attempts to open a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// This method first attempts to find the window in the existing windows collection.
        /// If not found, it uses the UI shell service to force-create and display the window.
        /// </remarks>
        /// <param name="objectKind">The tool window object kind GUID in string format.</param>
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
