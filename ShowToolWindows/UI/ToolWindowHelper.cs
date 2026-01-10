using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Debugger.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using System;
using System.Collections.Generic;
using System.Linq;

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

        public IEnumerable<Window> GetAllToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

#pragma warning disable VSTHRD010
            var allWindows = _dte.Windows
                .Cast<Window>()
                .Where(window => string.Equals(window.Kind, WindowKindConsts.ToolWindowKind, StringComparison.OrdinalIgnoreCase))
                .Where(window => !string.Equals(window.ObjectKind, EnvDTE.Constants.vsWindowKindMainWindow, StringComparison.OrdinalIgnoreCase))
                .ToList();
#pragma warning restore VSTHRD010

            System.Diagnostics.Debug.WriteLine($"=== Total Windows Found: {allWindows.Count} ===");
            foreach (var w in allWindows)
            {
                System.Diagnostics.Debug.WriteLine($"Caption: '{w.Caption}', Kind: '{w.Kind}', ObjectKind: '{w.ObjectKind}'");
            }

            return allWindows
                .ToList();
        }

        public IEnumerable<ToolWindowEntry> GetAllToolWindowEntries()
        {
            var toolWindows = GetAllToolWindows();
            return toolWindows.Select(window => new ToolWindowEntry(window)).ToList();
        }


        public void RestoreToolWindowsFromStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var objectKinds = stash.WindowObjectKinds;

            for (int i = 0; i < objectKinds.Length; i++)
            {
                var objectKind = objectKinds[i];

                TryOpenToolWindowByObjectKind(objectKind);
            }
        }


        public void CloseToolWindowsNotInStash(IEnumerable<ToolWindowEntry> toolWindows, ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var stashObjectKinds = new HashSet<string>(stash.WindowObjectKinds, StringComparer.OrdinalIgnoreCase);

            var windowsToClose = toolWindows
                .Where(entry => entry.IsVisible && !stashObjectKinds.Contains(entry.ObjectKind))
                .ToList();

            foreach (var entry in windowsToClose)
            {
                SetToolWindowVisibility(entry, false);
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

    }
}
