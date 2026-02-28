using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Provides helper methods for opening and managing Visual Studio tool windows via <see cref="DTE"/> and <see cref="IVsUIShell"/>.
    /// </summary>
    internal class ToolWindowHelper
    {
        private readonly DTE _dte;
        private readonly IVsUIShell _uiShell;
        private HashSet<string> _excludedWindowObjectKinds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

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
        /// Gets or sets the collection of window object kinds to exclude from operations.
        /// </summary>
        /// <remarks>
        /// Tool windows with object kinds in this collection will be excluded from:
        /// - GetAllToolWindows results
        /// - CloseToolWindowsNotInStash operations
        /// The setter sanitizes the input by removing null, empty, and whitespace-only entries.
        /// GUIDs are normalized to include enclosing curly braces.
        /// </remarks>
        public IEnumerable<string> ExcludedWindowObjectKinds
        {
            get
            {
                return _excludedWindowObjectKinds.ToList();
            }
            set
            {
                if (value == null || !value.Any())
                {
                    _excludedWindowObjectKinds.Clear();
                }
                else
                {
                    _excludedWindowObjectKinds = new HashSet<string>(
                        value.Where(s => !string.IsNullOrWhiteSpace(s))
                             .Select(s => ObjectKindHelper.NormalizeObjectKind(s.Trim())),
                        StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        /// <summary>
        /// Gets all tool windows as <see cref="ToolWindowEntry"/> objects.
        /// </summary>
        /// <remarks>
        /// Each returned window is wrapped in a <see cref="ToolWindowEntry"/> for easier state-based operations.
        /// </remarks>
        /// <returns>An enumerable collection of <see cref="ToolWindowEntry"/> objects representing all available tool windows.</returns>
        public IEnumerable<ToolWindowEntry> GetAllToolWindowEntries()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var toolWindows = GetAllToolWindows();
            return toolWindows.Select(window => new ToolWindowEntry(window)).ToList();
        }

        /// <summary>
        /// Restores tool windows from a stash by opening all windows specified in the stash.
        /// </summary>
        /// <remarks>
        /// Iterates the object kinds in <paramref name="stash"/> and calls <see cref="TryOpenToolWindowByObjectKind(string)"/>
        /// for each one, in stash order.
        /// </remarks>
        /// <param name="stash">The stash containing the object kinds of windows to restore.</param>
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
        /// Compares current tool windows to the object kinds in <paramref name="stash"/> and closes any window not present.
        /// Tool windows listed in <see cref="ExcludedWindowObjectKinds"/> are not included in the candidate set.
        /// </remarks>
        /// <param name="stash">The stash containing the object kinds of windows that should remain open.</param>
        public void CloseToolWindowsNotInStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var allToolWindowEntries = GetAllToolWindowEntries().ToList();

            var stashObjectKinds = new HashSet<string>(stash.WindowObjectKinds, StringComparer.OrdinalIgnoreCase);

            var toolWindowsToClose = allToolWindowEntries
                .Where(entry => !stashObjectKinds.Contains(entry.ObjectKind))
                .ToList();

            SetToolWindowsVisibility(toolWindowsToClose, false);
        }


        /// <summary>
        /// Sets the visibility state of a single tool window.
        /// </summary>
        /// <remarks>
        /// Verifies that a matching window exists in the current tool window set, then opens or closes it.
        /// Opening delegates to <see cref="TryOpenToolWindowByObjectKind(string)"/> and closing delegates to <c>TryCloseToolWindowByObjectKind</c>.
        /// </remarks>
        /// <param name="entry">The tool window entry to modify.</param>
        /// <param name="isVisible">Whether the tool window should be visible.</param>
        /// <returns><c>true</c> if the visibility was successfully changed; otherwise, <c>false</c>.</returns>
        public bool SetToolWindowVisibility(ToolWindowEntry entry, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var allToolWindowEntries = GetAllToolWindowEntries().ToList();

            var toolWindowToMutate = allToolWindowEntries
                .FirstOrDefault(toolWindowEntry => toolWindowEntry.ObjectKind == entry.ObjectKind);

            if (toolWindowToMutate == null)
            {
                return false;
            }

            if (isVisible)
            {
                return TryOpenToolWindowByObjectKind(entry.ObjectKind);
            }
            else
            {
                return TryCloseToolWindowByObjectKind(entry.ObjectKind);
            }
        }


        /// <summary>
        /// Sets the visibility state of multiple tool windows.
        /// </summary>
        /// <remarks>
        /// Iterates through <paramref name="toolWindows"/> and attempts to open or close each entry by object kind.
        /// This method does not pre-filter entries against the current window set.
        /// </remarks>
        /// <param name="toolWindows">The collection of tool window entries to modify.</param>
        /// <param name="isVisible">Whether the tool windows should be visible.</param>
        public void SetToolWindowsVisibility(IEnumerable<ToolWindowEntry> toolWindows, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (ToolWindowEntry entry in toolWindows)
            {
                if (isVisible)
                {
                    TryOpenToolWindowByObjectKind(entry.ObjectKind);
                }
                else
                {
                    TryCloseToolWindowByObjectKind(entry.ObjectKind);
                }
            }
        }


        /// <summary>
        /// Gets all tool windows available in the current Visual Studio instance.
        /// </summary>
        /// <remarks>
        /// This method filters the windows collection to include only tool windows,
        /// excluding the main window and other non-tool window types.
        /// Tool windows with object kinds in the ExcludedWindowObjectKinds collection will not be returned.
        /// </remarks>
        /// <returns>An enumerable collection of tool windows.</returns>
        private IEnumerable<Window> GetAllToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Cache COM properties to minimize COM interop calls
            var windowInfos = new List<(Window Window, string Kind, string ObjectKind, string Caption)>();

            foreach (Window window in _dte.Windows)
            {
                try
                {
                    string kind = window.Kind;
                    string objectKind = window.ObjectKind;
                    string caption = window.Caption;

                    windowInfos.Add((window, kind, objectKind, caption));
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error accessing window properties: {ex.Message}");
                }
            }

            // Exclude non-tool windows & VS main window
            var allWindows = windowInfos
                .Where(info => string.Equals(info.Kind, WindowKindConsts.ToolWindowKind, StringComparison.OrdinalIgnoreCase))
                .Where(info => !string.Equals(info.ObjectKind, EnvDTE.Constants.vsWindowKindMainWindow, StringComparison.OrdinalIgnoreCase))
                .Where(info => !_excludedWindowObjectKinds.Contains(info.ObjectKind))
                .ToList();

            return allWindows.Select(info => info.Window).ToList();
        }

        /// <summary>
        /// Attempts to open a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// First attempts to find the window in <see cref="DTE.Windows"/>.
        /// If found, the window is made visible and activated.
        /// If not found, it attempts force-create via <see cref="IVsUIShell.FindToolWindow(uint, ref Guid, out IVsWindowFrame)"/> and shows the frame.
        /// Exceptions are caught and written to debug output.
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
        /// Attempts to close a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// Attempts to find the window in <see cref="DTE.Windows"/> and set <c>Visible</c> to <c>false</c>.
        /// Exceptions are caught and written to debug output.
        /// </remarks>
        /// <param name="objectKind">The tool window object kind GUID in string format.</param>
        /// <returns><c>true</c> if the window is found and closed; otherwise, <c>false</c>.</returns>
        private bool TryCloseToolWindowByObjectKind(string objectKind)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var window = _dte.Windows.Item(objectKind);
                if (window != null)
                {
                    window.Visible = false;
                    return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Window '{objectKind}' not in collection. {ex.Message}");
            }

            return false;
        }
    }
}