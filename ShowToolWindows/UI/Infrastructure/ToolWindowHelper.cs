using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using ShowToolWindows.UI.Views.ToolWindows;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Provides helper methods for opening and managing Visual Studio tool windows.
    /// </summary>
    /// <remarks>
    /// This helper coordinates tool window operations by using <see cref="DTE"/> and <see cref="IVsUIShell"/>.
    /// </remarks>
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
        private ToolWindowHelper(DTE dte, IVsUIShell vsUiShell)
        {
            _dte = dte ?? throw new ArgumentNullException(nameof(dte));
            _uiShell = vsUiShell ?? throw new ArgumentNullException(nameof(vsUiShell));
        }

        /// <summary>
        /// Gets or sets the collection of window object kinds to exclude from operations.
        /// </summary>
        /// <value>A collection of tool window object kinds to exclude from helper operations.</value>
        /// <remarks>
        /// <para>Tool windows with object kinds in this collection are excluded from the following operations:</para>
        /// <list type="bullet">
        /// <item>
        /// <description><see cref="GetAllToolWindows"/> results.</description>
        /// </item>
        /// <item>
        /// <description><see cref="CloseToolWindowsNotInStash(ToolWindowStash)"/> operations.</description>
        /// </item>
        /// <item>
        /// <description><see cref="CloseVisibleToolWindows"/> operations.</description>
        /// </item>
        /// </list>
        /// <para>Values are sanitised before storage.</para>
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
                    _excludedWindowObjectKinds = value
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => ObjectKindHelper.NormalizeObjectKind(s.Trim()))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        /// <summary>
        /// Creates a new <see cref="ToolWindowHelper"/> instance with the default exclusions applied.
        /// </summary>
        /// <remarks>
        /// The created helper excludes <see cref="StashRestoreToolWindowsToolWindow.ToolWindowGuidString"/> by default.
        /// </remarks>
        /// <param name="dte">The Visual Studio automation object.</param>
        /// <param name="uiShell">The Visual Studio UI shell service.</param>
        /// <returns>A configured <see cref="ToolWindowHelper"/> instance.</returns>
        public static ToolWindowHelper Create(DTE dte, IVsUIShell uiShell)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var toolWindowHelper = new ToolWindowHelper(dte, uiShell)
            {
                ExcludedWindowObjectKinds = new HashSet<string> { StashRestoreToolWindowsToolWindow.ToolWindowGuidString }
            };

            return toolWindowHelper;
        }

        /// <summary>
        /// Gets all available tool windows as <see cref="ToolWindowEntry"/> objects.
        /// </summary>
        /// <remarks>
        /// Excluded tool windows are not included in the returned collection.
        /// </remarks>
        /// <returns>A read-only list of <see cref="ToolWindowEntry"/> objects representing all available tool windows.</returns>
        public IReadOnlyList<ToolWindowEntry> GetAllToolWindowEntries()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var toolWindows = GetAllToolWindows();
            return toolWindows.Select(window => new ToolWindowEntry(window)).ToList();
        }

        /// <summary>
        /// Restores the tool windows defined by the specified stash.
        /// </summary>
        /// <remarks>
        /// Only tool windows represented by <paramref name="stash"/> are restored.
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
        /// Tool windows not represented by <paramref name="stash"/> are closed.
        /// Tool windows listed in <see cref="ExcludedWindowObjectKinds"/> remain unaffected.
        /// </remarks>
        /// <param name="stash">The stash containing the object kinds of windows that should remain open.</param>
        public void CloseToolWindowsNotInStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var allToolWindowEntries = GetAllToolWindowEntries().ToList();

            var stashObjectKinds = stash.WindowObjectKinds.ToHashSet(StringComparer.OrdinalIgnoreCase);

            var toolWindowsToClose = allToolWindowEntries
                .Where(entry => !stashObjectKinds.Contains(entry.ObjectKind))
                .ToList();

            SetToolWindowsVisibility(toolWindowsToClose, false);
        }

        /// <summary>
        /// Closes all visible tool windows except those listed in <see cref="ExcludedWindowObjectKinds"/>.
        /// </summary>
        /// <remarks>
        /// Tool windows in <see cref="ExcludedWindowObjectKinds"/> remain open.
        /// </remarks>
        /// <returns>The number of tool windows that were successfully closed.</returns>
        public int CloseVisibleToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var candidates = GetAllToolWindows();

            if (ExcludedWindowObjectKinds.Any())
            {
#pragma warning disable VSTHRD010
                candidates = candidates.Where(w => !ExcludedWindowObjectKinds.Contains(w.ObjectKind));
#pragma warning restore VSTHRD010
            }

            return WindowHelper.CloseWindows(candidates);
        }

        /// <summary>
        /// Sets the visibility state of a single tool window.
        /// </summary>
        /// <remarks>
        /// If the specified tool window cannot be found, the method returns <c>false</c>.
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
        /// Each tool window in <paramref name="toolWindows"/> is processed independently.
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
        /// Attempts to open a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// This method returns <c>true</c> when the requested tool window is available for display.
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
        /// Gets all tool windows available in the current Visual Studio instance.
        /// </summary>
        /// <remarks>
        /// Excluded tool windows are not included in the returned collection.
        /// </remarks>
        /// <returns>A collection of available tool windows.</returns>
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
        /// Attempts to close a tool window by its object kind GUID.
        /// </summary>
        /// <remarks>
        /// This method returns <c>true</c> when the requested tool window is found and closed.
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