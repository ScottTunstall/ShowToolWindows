using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using ShowToolWindows.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace ShowToolWindows.UI.ToolWindows
{
    /// <summary>
    /// User control that provides an interface for managing Visual Studio tool windows.
    /// Supports selecting and stashing tool window configurations.
    /// </summary>
    public partial class ToggleToolWindowsControl : UserControl, INotifyPropertyChanged
    {
        private AsyncPackage _package;
        private DTE _dte;
        private IVsUIShell _uiShell;
        private StashSettingsService _stashService;
        private ToolWindowHelper _toolWindowHelper;
        private bool _isInitialised;
        private bool _hasSelectedItems;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleToolWindowsControl"/> class.
        /// </summary>
        public ToggleToolWindowsControl()
        {
            InitializeComponent();
            DataContext = this;

            RefreshCommand = new RelayCommand(ExecuteRefresh);
            DropStashCommand = new RelayCommand(parameter => ExecuteDropStash());
            CheckAllCommand = new RelayCommand(ExecuteCheckAll);

            // Subscribe to collection events early to ensure we capture all changes,
            // including those made during async initialization.
            Stashes.CollectionChanged += Stashes_CollectionChanged;
            ToolWindows.CollectionChanged += ToolWindows_CollectionChanged;
        }

        /// <summary>
        /// Gets the command for refreshing the tool windows list.
        /// Bound to the F5 keybinding for user convenience.
        /// </summary>
        public ICommand RefreshCommand { get; }

        /// <summary>
        /// Gets the command for dropping a stash.
        /// Bound to the Delete key on the StashListBox.
        /// </summary>
        public ICommand DropStashCommand { get; }

        /// <summary>
        /// Gets the command for checking all tool windows.
        /// Bound to the Ctrl+A keybinding for user convenience.
        /// </summary>
        public ICommand CheckAllCommand { get; }

        /// <summary>
        /// Gets the collection of stashed tool window snapshots.
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global - needs to be public so Xaml can bind to it
        public ObservableCollection<ToolWindowStash> Stashes
        {
            get;
        } = new ObservableCollection<ToolWindowStash>();

        /// <summary>
        /// Gets the collection of tool windows displayed in the UI.
        /// </summary>
        // ReSharper disable once MemberCanBePrivate.Global - needs to be public so Xaml can bind to it
        public ObservableCollection<ToolWindowEntry> ToolWindows
        {
            get;
        } = new ObservableCollection<ToolWindowEntry>();

        /// <summary>
        /// Gets the header text for the stashes expander, displaying the count of available stashes.
        /// </summary>
        public string StashesHeader => $"Stashes ({Stashes.Count})";

        /// <summary>
        /// Gets a value indicating whether there are any stashes available.
        /// </summary>
        public bool HaveStashes => Stashes.Count > 0;

        /// <summary>
        /// Gets a value indicating whether stash modification operations (pop, drop) can be performed.
        /// Requires both initialization and at least one stash to be available.
        /// </summary>
        public bool CanMutateStash => IsInitialised && HaveStashes;

        /// <summary>
        /// Gets a value indicating whether the stash button should be enabled.
        /// Requires both initialization and at least one checked tool window.
        /// </summary>
        public bool CanStashSelected => IsInitialised && HaveSelectedItems;

        /// <summary>
        /// Gets or sets a value indicating whether any items are checked in the tool windows list.
        /// </summary>
        public bool HaveSelectedItems
        {
            get => _hasSelectedItems;
            private set
            {
                if (_hasSelectedItems != value)
                {
                    _hasSelectedItems = value;
                    OnPropertyChanged(nameof(HaveSelectedItems));
                    OnPropertyChanged(nameof(CanStashSelected));
                }
            }
        }


        /// <summary>
        /// Gets a value indicating whether the control has been initialized with Visual Studio services.
        /// </summary>
        public bool IsInitialised
        {
            get => _isInitialised;
            private set
            {
                if (_isInitialised != value)
                {
                    _isInitialised = value;
                    OnPropertyChanged(nameof(IsInitialised));
                    OnPropertyChanged(nameof(CanMutateStash));
                    OnPropertyChanged(nameof(CanStashSelected));
                }
            }
        }

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Initializes the control with Visual Studio services asynchronously.
        /// </summary>
        /// <param name="package">The owning package.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="package"/> is null.</exception>
        /// <exception cref="InvalidOperationException">Thrown when required Visual Studio services cannot be obtained.</exception>
        public async Task InitializeAsync(AsyncPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            if (_package != null)
            {
                RefreshToolWindows();
                return;
            }

            _package = package;

            var dteService = await package.GetServiceAsync(typeof(DTE));
            _dte = dteService as DTE ?? throw new InvalidOperationException("Failed to get DTE service.");

            _uiShell = await _package.GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell ?? throw new InvalidOperationException("Failed to get IVsUIShell service.");

            _stashService = new StashSettingsService(_package);
            _toolWindowHelper = new ToolWindowHelper(_dte, _uiShell)
            {
                // We do not want to allow stashing or toggling this tool window itself
                ExcludedWindowObjectKinds = new HashSet<string>() { StashRestoreToolWindowsToolWindow.ToolWindowGuidString }
            };
            LoadAllToolWindowStashes();

            IsInitialised = true;
            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        /// <summary>
        /// Handles the CollectionChanged event of the Stashes collection to update related UI properties.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="NotifyCollectionChangedEventArgs"/> instance containing the event data.</param>
        private void Stashes_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(StashesHeader));
            OnPropertyChanged(nameof(HaveStashes));
            OnPropertyChanged(nameof(CanMutateStash));
        }

        /// <summary>
        /// Handles the CollectionChanged event of the ToolWindows collection to subscribe to item property changes.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="NotifyCollectionChangedEventArgs"/> instance containing the event data.</param>
        private void ToolWindows_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (e.OldItems != null)
            {
                foreach (ToolWindowEntry item in e.OldItems)
                {
                    item.PropertyChanged -= ToolWindowEntry_PropertyChanged;
                }
            }

            if (e.NewItems != null)
            {
                foreach (ToolWindowEntry item in e.NewItems)
                {
                    item.PropertyChanged += ToolWindowEntry_PropertyChanged;
                }
            }

            UpdateHaveSelectedItems();
        }

        /// <summary>
        /// Handles property changes on ToolWindowEntry items to update the HaveSelectedItems property.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="PropertyChangedEventArgs"/> instance containing the event data.</param>
        private void ToolWindowEntry_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(ToolWindowEntry.IsSelected))
            {
                UpdateHaveSelectedItems();
            }
        }

        /// <summary>
        /// Updates the HaveSelectedItems property based on the current checked state of tool windows.
        /// </summary>
        private void UpdateHaveSelectedItems()
        {
            HaveSelectedItems = ToolWindows.Any(w => w.IsSelected);
        }

        /// <summary>
        /// Raises the PropertyChanged event for the specified property name.
        /// </summary>
        /// <param name="propertyName">Name of the property that changed.</param>
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Handles the Refresh button click to reload the tool windows list.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void RefreshButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteRefreshToolWindows();
        }

        /// <summary>
        /// Handles the Stash button click to save the currently checked tool windows to a new stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void StashButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteStashSelectedToolWindows();
        }

        /// <summary>
        /// Handles the Pop/Merge button click to restore tool windows from the top stash while keeping current windows open.
        /// The stash is removed after restoration.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void PopMergeButton_Click(object sender, RoutedEventArgs e)
        {
            ExecutePopAndMergeToolWindowsFromStash();
        }

        /// <summary>
        /// Handles the Pop (Abs) button click to restore tool windows from the top stash in absolute mode.
        /// Closes all windows not in the stash before restoring. The stash is removed after restoration.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void PopAbsoluteButton_Click(object sender, RoutedEventArgs e)
        {
            ExecutePopToolWindowsFromStashAbsolute();
        }

        /// <summary>
        /// Handles the Drop All button click to clear all stashes after user confirmation.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DropAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDropAllStashes();
        }

        /// <summary>
        /// Handles double-click on a stash item to restore the tool windows from that stash.
        /// Left-click restores in merge mode. Right-click restores in absolute mode.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="MouseButtonEventArgs"/> instance containing the event data.</param>
        private void StashListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }

            if (e.LeftButton == MouseButtonState.Pressed)
            {
                ExecuteRestoreToolWindowsFromStash(stash);
            }
        }

        /// <summary>
        /// Handles right-click on a stash item to select it before the context menu appears.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="MouseButtonEventArgs"/> instance containing the event data.</param>
        private void StashListBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var listBoxItem = FindAncestor<ListBoxItem>((DependencyObject)e.OriginalSource);

            if (listBoxItem != null)
            {
                listBoxItem.IsSelected = true;
                listBoxItem.Focus();
            }
        }

        /// <summary>
        /// Handles the Apply (Absolute) context menu click to restore tool windows from the selected stash in absolute mode.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ApplyStashAbsolute_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }

            ExecuteRestoreToolWindowsFromStashAbsolute(stash);
        }

        /// <summary>
        /// Handles the Apply (Merge) context menu click to restore tool windows from the selected stash in merge mode.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ApplyStashMerge_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }

            ExecuteRestoreToolWindowsFromStash(stash);
        }

        /// <summary>
        /// Handles the Hide all visible context menu click to hide all tool windows referenced in the selected stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HideAllVisible_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }

            ExecuteHideAllVisibleInStash(stash);
        }

        /// <summary>
        /// Handles the Drop context menu click to remove the selected stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DropStash_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDropStash();
        }

        /// <summary>
        /// Executes the refresh command to reload the tool windows list.
        /// </summary>
        /// <param name="parameter">The command parameter (unused).</param>
        private void ExecuteRefresh(object parameter)
        {
            RefreshToolWindows();
        }

        /// <summary>
        /// Executes the check all command to select all tool windows.
        /// </summary>
        /// <param name="parameter">The command parameter (unused).</param>
        private void ExecuteCheckAll(object parameter)
        {
            if (!IsInitialised)
            {
                return;
            }

            foreach (var window in ToolWindows)
            {
                window.IsSelected = true;
            }
        }

        /// <summary>
        /// Executes the refresh operation and displays a status bar notification.
        /// </summary>
        private void ExecuteRefreshToolWindows()
        {
            RefreshToolWindows();

            StatusBarHelper.ShowStatusBarNotification("Tool windows refreshed.");
        }

        /// <summary>
        /// Creates a new stash from the currently checked tool windows and displays a status bar notification.
        /// </summary>
        private void ExecuteStashSelectedToolWindows()
        {
            if (!IsInitialised)
            {
                return;
            }

            int selectedCount = ToolWindows.Count(w => w.IsSelected);

            if (selectedCount == 0)
            {
                StatusBarHelper.ShowStatusBarNotification("No tool windows checked to stash. Check items in the list first.");
                return;
            }

            StashSelectedToolWindows();

            StatusBarHelper.ShowStatusBarNotification($"{selectedCount} checked tool window(s) stashed.");
        }

        /// <summary>
        /// Restores tool windows from the top stash in merge mode (keeps existing windows open) and displays a status bar notification.
        /// </summary>
        private void ExecutePopAndMergeToolWindowsFromStash()
        {
            if (Stashes.Count == 0)
            {
                return;
            }

            PopToolWindowsFromStash();

            StatusBarHelper.ShowStatusBarNotification("Tool windows merged from stash.");
        }

        /// <summary>
        /// Restores tool windows from the top stash in absolute mode (closes windows not in the stash) and displays a status bar notification.
        /// </summary>
        private void ExecutePopToolWindowsFromStashAbsolute()
        {
            if (Stashes.Count == 0)
            {
                return;
            }

            PopAbsToolWindowsFromStash();

            StatusBarHelper.ShowStatusBarNotification("Tool windows replaced by stash.");
        }

        /// <summary>
        /// Executes the drop stash command to remove the selected stash from the collection after user confirmation.
        /// </summary>
        private void ExecuteDropStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            if (!(this.StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }

            Guid clsid = Guid.Empty;
            _uiShell.ShowMessageBox(
                0,
                ref clsid,
                "Drop Stash",
                "Are you sure you wish to drop this stash?",
                null,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND,
                OLEMSGICON.OLEMSGICON_QUERY,
                0,
                out int result);

            if (result != (int)VSConstants.MessageBoxResult.IDYES)
            {
                return;
            }

            DropStash(stash);
            StatusBarHelper.ShowStatusBarNotification("Stash dropped.");
        }

        /// <summary>
        /// Clears all stashes after user confirmation and displays a status bar notification.
        /// </summary>
        private void ExecuteDropAllStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            Guid clsid = Guid.Empty;
            _uiShell.ShowMessageBox(
                0,
                ref clsid,
                "Drop All Stashes",
                "Are you sure you wish to drop all stashes? This cannot be undone.",
                null,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNO,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND,
                OLEMSGICON.OLEMSGICON_QUERY,
                0,
                out int result);

            if (result != (int)VSConstants.MessageBoxResult.IDYES)
            {
                return;
            }

            DropAllStashes();
            StatusBarHelper.ShowStatusBarNotification("All tool window stashes dropped.");
        }

        /// <summary>
        /// Restores tool windows from a specific stash in merge mode (keeps existing windows open).
        /// Does not remove the stash.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to restore.</param>
        private void ExecuteRestoreToolWindowsFromStash(ToolWindowStash stash)
        {
            RestoreToolWindowsFromStash(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows restored from stash.");
        }

        /// <summary>
        /// Restores tool windows from a specific stash in absolute mode (closes windows not in the stash).
        /// Does not remove the stash.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to restore.</param>
        private void ExecuteRestoreToolWindowsFromStashAbsolute(ToolWindowStash stash)
        {
            RestoreToolWindowsFromStashAbsolute(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows restored from stash.");
        }

        /// <summary>
        /// Hides all visible tool windows that are referenced in the specified stash.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to hide.</param>
        private void ExecuteHideAllVisibleInStash(ToolWindowStash stash)
        {
            HideAllVisibleInStash(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows in stash hidden.");
        }

        /// <summary>
        /// Creates a new stash from the currently checked tool windows.
        /// The stash is added to the top of the stash collection and persisted to settings.
        /// </summary>
        private void StashSelectedToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            var windowsToStash = ToolWindows
                .Where(w => w.IsSelected)
                .OrderBy(entry => entry.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (windowsToStash.Count == 0)
            {
                return;
            }

            var captions = new List<string>();
            var objectKinds = new List<string>();

            foreach (var entry in windowsToStash)
            {
                captions.Add(entry.Caption);
                objectKinds.Add(entry.ObjectKind);
            }

            var stash = new ToolWindowStash
            {
                WindowCaptions = captions.ToArray(),
                WindowObjectKinds = objectKinds.ToArray(),
                CreatedAt = DateTimeOffset.UtcNow
            };

            Stashes.Insert(0, stash);
            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Restores tool windows from the top stash in merge mode and removes the stash.
        /// Does not close windows that are not in the stash.
        /// </summary>
        private void PopToolWindowsFromStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            if (Stashes.Count == 0)
            {
                return;
            }

            var stash = Stashes[0];
            RestoreToolWindowsFromStash(stash);
            Stashes.RemoveAt(0);
            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Restores tool windows from the top stash in absolute mode and removes the stash.
        /// Closes all windows not in the stash before restoring.
        /// </summary>
        private void PopAbsToolWindowsFromStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            if (Stashes.Count == 0)
            {
                return;
            }

            var stash = Stashes[0];

            RestoreToolWindowsFromStashAbsolute(stash);
            Stashes.RemoveAt(0);
            SaveAllToolWindowStashes();
        }

        private void DropStash(ToolWindowStash stash)
        {
            Stashes.Remove(stash);
            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Clears all stashes and saves the updated stash list.
        /// </summary>
        private void DropAllStashes()
        {
            Stashes.Clear();
            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Loads all saved tool window stashes from persistent settings.
        /// </summary>
        private void LoadAllToolWindowStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Stashes.Clear();
            var loadedStashes = _stashService.LoadStashes();

            foreach (var stash in loadedStashes)
            {
                Stashes.Add(stash);
            }
        }

        /// <summary>
        /// Saves all current tool window stashes to persistent settings.
        /// </summary>
        private void SaveAllToolWindowStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            var stashList = new List<ToolWindowStash>(Stashes);
            _stashService.SaveStashes(stashList);
        }

        /// <summary>
        /// Restores (shows) the tool windows specified in the given stash in merge mode and refreshes the tool windows list.
        /// Does not close windows that are not in the stash.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to restore.</param>
        private void RestoreToolWindowsFromStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            _toolWindowHelper.RestoreToolWindowsFromStash(stash);

            RefreshToolWindows();
        }

        /// <summary>
        /// Restores tool windows from the given stash in absolute mode and refreshes the tool windows list.
        /// Closes all windows not in the stash before restoring.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to restore.</param>
        private void RestoreToolWindowsFromStashAbsolute(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            _toolWindowHelper.CloseToolWindowsNotInStash(stash);
            _toolWindowHelper.RestoreToolWindowsFromStash(stash);

            RefreshToolWindows();
        }

        /// <summary>
        /// Hides all visible tool windows that are referenced in the specified stash.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to hide.</param>
        private void HideAllVisibleInStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            var allToolWindowEntries = _toolWindowHelper.GetAllToolWindowEntries().ToList();
            var stashObjectKinds = new HashSet<string>(stash.WindowObjectKinds, StringComparer.OrdinalIgnoreCase);

            var toolWindowsToHide = allToolWindowEntries
                .Where(entry => stashObjectKinds.Contains(entry.ObjectKind))
                .ToList();

            _toolWindowHelper.SetToolWindowsVisibility(toolWindowsToHide, false);

            RefreshToolWindows();
        }

        /// <summary>
        /// Refreshes the tool windows list by querying Visual Studio for all available tool windows.
        /// Filters out unsupported windows and populates the list with unchecked items.
        /// </summary>
        private void RefreshToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            ToolWindows.Clear();

            var allToolWindowEntries = _toolWindowHelper.GetAllToolWindowEntries();
            var supportedToolWindowEntries = allToolWindowEntries
                .Where(entry => IsSupportedToolWindow(entry))
                .OrderBy(entry => entry.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            foreach (var entry in supportedToolWindowEntries)
            {
                ToolWindows.Add(entry);
            }
        }

        /// <summary>
        /// Finds an ancestor of a specific type in the visual tree.
        /// </summary>
        /// <typeparam name="T">The type of ancestor to find.</typeparam>
        /// <param name="current">The starting dependency object.</param>
        /// <returns>The ancestor of type T, or null if not found.</returns>
        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T ancestor)
                {
                    return ancestor;
                }
                current = System.Windows.Media.VisualTreeHelper.GetParent(current);
            }
            return null;
        }

        /// <summary>
        /// Determines whether a tool window is supported for management by this control.
        /// Excludes tool windows with object kinds in the ExcludedWindowObjectKinds collection.
        /// </summary>
        /// <param name="windowEntry">The tool window entry to check.</param>
        /// <returns>True if the tool window is supported; otherwise, false.</returns>
        private bool IsSupportedToolWindow(ToolWindowEntry windowEntry)
        {
            if (windowEntry == null)
            {
                return false;
            }

            string objectKindNormalized = windowEntry.ObjectKind.Trim('{', '}');
            if (string.Equals(objectKindNormalized, StashRestoreToolWindowsToolWindow.ToolWindowGuidString, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(windowEntry.Caption, StashRestoreToolWindowsToolWindow.ToolWindowTitle, StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }

#pragma warning restore VSTHRD010
    }
}