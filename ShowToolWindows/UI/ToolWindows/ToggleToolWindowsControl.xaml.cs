using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using ShowToolWindows.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Window = EnvDTE.Window;

namespace ShowToolWindows.UI.ToolWindows
{
    /// <summary>
    /// Interaction logic for CloseToolWindowsControl.
    /// </summary>
    public partial class ToggleToolWindowsControl : UserControl
    {
        private AsyncPackage _package;
        private DTE _dte;
        private StashSettingsService _stashService;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleToolWindowsControl"/> class.
        /// </summary>
        public ToggleToolWindowsControl()
        {
            InitializeComponent();
            DataContext = this;

            RefreshCommand = new RelayCommand(ExecuteRefresh);
        }

        /// <summary>
        /// Gets the command for refreshing the tool windows list (F5).
        /// </summary>
        public ICommand RefreshCommand { get; }

        /// <summary>
        /// Gets the collection of tool windows displayed in the UI.
        /// </summary>
        public ObservableCollection<ToolWindowEntry> ToolWindows
        {
            get;
        } = new ObservableCollection<ToolWindowEntry>();

        /// <summary>
        /// Gets the collection of stashed tool window snapshots.
        /// </summary>
        public ObservableCollection<ToolWindowStash> Stashes
        {
            get;
        } = new ObservableCollection<ToolWindowStash>();

        /// <summary>
        /// Initializes the control with Visual Studio services.
        /// </summary>
        /// <param name="package">The owning package.</param>
        public async Task InitializeAsync(AsyncPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            if (_package != null)
            {
                RefreshToolWindows();
                return;
            }

            _package = package;

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            object dteService = await package.GetServiceAsync(typeof(DTE));
            _dte = dteService as DTE;

            _stashService = new StashSettingsService(_package);
            LoadToolWindowStashes();

            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        /// <summary>
        /// Handles the Checked event to show a tool window.
        /// </summary>
        private void WindowCheckbox_Checked(object sender, RoutedEventArgs e)
        {
            SetToolWindowVisibility(sender, true);
        }

        /// <summary>
        /// Handles the Unchecked event to hide a tool window.
        /// </summary>
        private void WindowCheckbox_Unchecked(object sender, RoutedEventArgs e)
        {
            SetToolWindowVisibility(sender, false);
        }

        /// <summary>
        /// Handles the Refresh button click to reload the tool windows list.
        /// </summary>
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            RefreshToolWindows();
        }

        /// <summary>
        /// Handles the Show All button click to display all tool windows.
        /// </summary>
        private void ShowAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetAllToolWindowsVisibility(true);
        }

        /// <summary>
        /// Handles the Hide All button click to hide all tool windows.
        /// </summary>
        private void HideAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetAllToolWindowsVisibility(false);
        }

        /// <summary>
        /// Handles the Stash button click to save current visible tool windows.
        /// </summary>
        private void StashButton_Click(object sender, RoutedEventArgs e)
        {
            StashOpenToolWindows();
        }

        /// <summary>
        /// Handles the Pop button click to restore and remove the top stashed item.
        /// </summary>
        private void PopButton_Click(object sender, RoutedEventArgs e)
        {
            PopToolWindowsFromStash();
        }


        /// <summary>
        /// Handles double-click on a stash to restore those tool windows.
        /// </summary>
        private void StashListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (StashListBox.SelectedItem is ToolWindowStash stash)
            {
                RestoreToolWindowStash(stash);
            }
        }

        private void ExecuteRefresh(object parameter)
        {
            RefreshToolWindows();
        }

        private void RefreshToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ToolWindows.Clear();

            if (_dte == null)
            {
                return;
            }

            var allWindows = _dte.Windows.Cast<Window>().ToList();
            
            System.Diagnostics.Debug.WriteLine($"=== Total Windows Found: {allWindows.Count} ===");
            foreach (var w in allWindows)
            {
                System.Diagnostics.Debug.WriteLine($"Caption: '{w.Caption}', Kind: '{w.Kind}', ObjectKind: '{w.ObjectKind}'");
            }

            var windows = allWindows
                .Where(IsSupportedToolWindow)
                .OrderBy(w => w.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            foreach (Window window in windows)
            {
                ToolWindows.Add(new ToolWindowEntry(window));
            }
        }

        private static bool IsSupportedToolWindow(Window window)
        {
            if (window == null)
            {
                return false;
            }

            if (!string.Equals(window.Kind, WindowKindConsts.ToolWindowKind, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(window.ObjectKind, EnvDTE.Constants.vsWindowKindMainWindow, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string objectKindNormalized = window.ObjectKind.Trim('{', '}');
            if (string.Equals(objectKindNormalized, ToggleToolWindowsToolWindow.ToolWindowGuidString,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(window.Caption, ToggleToolWindowsToolWindow.ToolWindowTitle, StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }

        private void SetToolWindowVisibility(object sender, bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            FrameworkElement frameworkElement = sender as FrameworkElement;
            if (!(frameworkElement?.DataContext is ToolWindowEntry entry))
            {
                return;
            }

            entry.SetVisibility(isVisible);
            entry.Synchronize();
        }

        private void SetAllToolWindowsVisibility(bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (ToolWindowEntry entry in ToolWindows)
            {
                entry.SetVisibility(isVisible);
                entry.Synchronize();
            }
        }

        private void StashOpenToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var visibleWindows = _dte.Windows.Cast<Window>()
                .Where(w => w.Visible)
                .Where(IsSupportedToolWindow)
                .OrderBy(w => w.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (visibleWindows.Count == 0)
            {
                return;
            }

            var captions = new List<string>();
            var objectKinds = new List<string>();

            foreach (var window in visibleWindows)
            {
                captions.Add(window.Caption);
                objectKinds.Add(window.ObjectKind);
            }

            var stash = new ToolWindowStash
            {
                WindowCaptions = captions.ToArray(),
                WindowObjectKinds = objectKinds.ToArray(),
                CreatedAt = DateTime.Now
            };

            Stashes.Insert(0, stash);
            SaveToolWindowStashes();
        }

        private void PopToolWindowsFromStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (Stashes.Count == 0)
            {
                return;
            }

            var stash = Stashes[0];
            RestoreToolWindowStash(stash);
            Stashes.RemoveAt(0);
            SaveToolWindowStashes();
        }

        private void LoadToolWindowStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_stashService == null)
            {
                return;
            }

            Stashes.Clear();
            var loadedStashes = _stashService.LoadStashes();
           
            foreach (var stash in loadedStashes)
            {
                Stashes.Add(stash);
            }
        }

        private void SaveToolWindowStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (_stashService == null)
            {
                return;
            }

            var stashList = new List<ToolWindowStash>(Stashes);
            _stashService.SaveStashes(stashList);
        }

        private void RestoreToolWindowStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var captions = stash.WindowCaptions;

            var objectKinds = stash.WindowObjectKinds;

            var objectKindsAsHashSet = new HashSet<string>(objectKinds);

            var allWindows = _dte.Windows.Cast<Window>()
                .Where(IsSupportedToolWindow)
                .ToList();

            foreach (var window in allWindows)
            {
                var shouldBeVisible = objectKindsAsHashSet.Contains(window.ObjectKind);

                try
                {
                    if (shouldBeVisible && !window.Visible)
                    {
                        window.Visible = true;
                        window.Activate();
                    }
                    else if (!shouldBeVisible && window.Visible)
                    {
                        window.Visible = false;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not set visibility for window '{window.Caption}': {ex.Message}");
                }
            }

            for (int i = 0; i < objectKinds.Length; i++)
            {
                var objectKind = objectKinds[i];
                var caption = i < captions.Length ? captions[i] : "Unknown";

                if (string.IsNullOrWhiteSpace(objectKind))
                {
                    continue;
                }

                var windowExists = allWindows.Any(w => 
                    string.Equals(w.ObjectKind, objectKind, StringComparison.OrdinalIgnoreCase));

                if (!windowExists)
                {
                    TryOpenWindowByObjectKind(objectKind, caption);
                }
            }

            RefreshToolWindows();
        }

        private void TryOpenWindowByObjectKind(string objectKind, string caption)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var window = _dte.Windows.Item(objectKind);
                if (window != null)
                {
                    window.Visible = true;
                    window.Activate();
                    System.Diagnostics.Debug.WriteLine($"Opened window '{caption}' using ObjectKind '{objectKind}'");
                    return;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Window '{caption}' not in collection, will try to create it: {ex.Message}");
            }

            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                try
                {
                    IVsUIShell uiShell = await _package.GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell;
                    if (uiShell == null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not get IVsUIShell to open window '{caption}'");
                        return;
                    }

                    Guid toolWindowGuid = new Guid(objectKind);
                    int hr = uiShell.FindToolWindow((uint)__VSFINDTOOLWIN.FTW_fForceCreate, ref toolWindowGuid, out IVsWindowFrame frame);
                    if (ErrorHandler.Succeeded(hr) && frame != null)
                    {
                        frame.Show();
                        System.Diagnostics.Debug.WriteLine($"Opened window '{caption}' using IVsUIShell and ObjectKind '{objectKind}'");
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"FindToolWindow failed for '{caption}' (ObjectKind '{objectKind}'), hr={hr}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not open window '{caption}' by ObjectKind '{objectKind}' via IVsUIShell: {ex.Message}");
                }
            });
        }

#pragma warning restore VSTHRD010
    }
}