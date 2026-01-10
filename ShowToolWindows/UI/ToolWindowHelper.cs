using EnvDTE;
using Microsoft.VisualStudio;
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

        /// <summary>
        /// Gets all tool windows available in the current Visual Studio instance.
        /// </summary>
        /// <remarks>
        /// This method filters the windows collection to include only tool windows,
        /// excluding the main window and other non-tool window types.
        /// </remarks>
        /// <returns>An enumerable collection of tool windows.</returns>
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

        /// <summary>
        /// Gets all tool windows as <see cref="ToolWindowEntry"/> objects.
        /// </summary>
        /// <remarks>
        /// This method wraps each tool window in a <see cref="ToolWindowEntry"/> object
        /// which provides a convenient interface for managing tool window state.
        /// </remarks>
        /// <returns>An enumerable collection of <see cref="ToolWindowEntry"/> objects representing all available tool windows.</returns>
        public IEnumerable<ToolWindowEntry> GetAllToolWindowEntries()
        {
            var toolWindows = GetAllToolWindows();
            return toolWindows.Select(window => new ToolWindowEntry(window)).ToList();
        }

        /// <summary>
        /// Restores tool windows from a stash by opening all windows specified in the stash.
        /// </summary>
        /// <remarks>
        /// This method iterates through all object kinds stored in the stash and attempts
        /// to open each corresponding tool window. Windows are opened in the order they
        /// appear in the stash.
        /// </remarks>
        /// <param name="stash">The stash containing the object kinds of windows to restore.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="stash"/> is null.</exception>
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

        /// <summary>
        /// Closes all tool windows that are not present in the specified stash.
        /// </summary>
        /// <remarks>
        /// This method compares the provided collection of tool windows against the stash
        /// and closes any windows whose object kind is not found in the stash.
        /// This is useful for restoring an absolute window state from a stash.
        /// </remarks>
        /// <param name="toolWindows">The collection of currently available tool windows to check.</param>
        /// <param name="stash">The stash containing the object kinds of windows that should remain open.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="toolWindows"/> or <paramref name="stash"/> is null.</exception>
        public void CloseToolWindowsNotInStash(IEnumerable<ToolWindowEntry> toolWindows, ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var stashObjectKinds = new HashSet<string>(stash.WindowObjectKinds, StringComparer.OrdinalIgnoreCase);

            var toolWindowsToClose = toolWindows
                .Where(entry => !stashObjectKinds.Contains(entry.ObjectKind))
                .ToList();

            SetToolWindowsVisibility(toolWindowsToClose, false);
        }

        /// <summary>
        /// Attempts to open a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// This method first attempts to find the window in the existing windows collection.
        /// If found, it makes the window visible and activates it.
        /// If not found, it uses the UI shell service to force-create and display the window.
        /// </remarks>
        /// <param name="objectKind">The tool window object kind GUID in string format.</param>
        /// <returns><c>true</c> if the window is opened or already available; otherwise, <c>false</c>.</returns>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="objectKind"/> is null.</exception>
        /// <exception cref="FormatException">Thrown when <paramref name="objectKind"/> is not a valid GUID format.</exception>
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
        }

        /// <summary>
        /// Sets the visibility state of a single tool window.
        /// </summary>
        /// <remarks>
        /// This method updates the visibility state of the specified tool window entry
        /// and synchronizes the changes with the underlying Visual Studio window object.
        /// </remarks>
        /// <param name="entry">The tool window entry to modify.</param>
        /// <param name="isVisible">Whether the tool window should be visible.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="entry"/> is null.</exception>
        public void SetToolWindowVisibility(ToolWindowEntry entry, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            entry.SetVisibility(isVisible);
            entry.Synchronize();
        }

        /// <summary>
        /// Sets the visibility state of multiple tool windows.
        /// </summary>
        /// <remarks>
        /// This method iterates through the provided collection and updates the visibility
        /// state of each tool window entry, synchronizing changes with the underlying
        /// Visual Studio window objects.
        /// </remarks>
        /// <param name="toolWindows">The collection of tool window entries to modify.</param>
        /// <param name="isVisible">Whether the tool windows should be visible.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="toolWindows"/> is null.</exception>
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