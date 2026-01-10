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
    /// Supports showing, hiding, and stashing tool window configurations.
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

            Loaded += ToggleToolWindowsControl_Loaded;
            Unloaded += ToggleToolWindowsControl_Unloaded;
        }


        /// <summary>
        /// Gets the command for refreshing the tool windows list.
        /// Bound to the F5 keybinding for user convenience.
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
        /// Requires both initialization and at least one selected tool window.
        /// </summary>
        public bool CanStashSelected => IsInitialised && HaveSelectedItems;

        /// <summary>
        /// Gets or sets a value indicating whether any items are selected in the tool windows list box.
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
        /// Gets the collection of stashed tool window snapshots.
        /// </summary>
        public ObservableCollection<ToolWindowStash> Stashes
        {
            get;
        } = new ObservableCollection<ToolWindowStash>();

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

            object dteService = await package.GetServiceAsync(typeof(DTE));
            _dte = dteService as DTE ?? throw new InvalidOperationException("Failed to get DTE service.");

            _uiShell = await _package.GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell ?? throw new InvalidOperationException("Failed to get IVsUIShell service.");

            _stashService = new StashSettingsService(_package);
            _toolWindowHelper = new ToolWindowHelper(_dte, _uiShell)
            {
                // We do not want to allow stashing or toggling this tool window itself
                ExcludedWindowObjectKinds = new HashSet<string>() { ToggleToolWindowsToolWindow.ToolWindowGuidString }
            };
            LoadAllToolWindowStashes();

            IsInitialised = true;
            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        /// <summary>
        /// Handles the Loaded event of the control to subscribe to collection and selection change events.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ToggleToolWindowsControl_Loaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged += Stashes_CollectionChanged;
            ToolWindowsListBox.SelectionChanged += ToolWindowsListBox_SelectionChanged;
        }

        /// <summary>
        /// Handles the Unloaded event of the control to unsubscribe from collection and selection change events.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ToggleToolWindowsControl_Unloaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged -= Stashes_CollectionChanged;
            ToolWindowsListBox.SelectionChanged -= ToolWindowsListBox_SelectionChanged;
        }

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
        /// Handles the SelectionChanged event of the ToolWindowsListBox to update the HasSelectedItems property.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="SelectionChangedEventArgs"/> instance containing the event data.</param>
        private void ToolWindowsListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            HaveSelectedItems = ToolWindowsListBox.SelectedItems.Count > 0;
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
        /// Handles the Checked event of a tool window checkbox to show the corresponding tool window.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void WindowCheckbox_Checked(object sender, RoutedEventArgs e)
        {
            SelectCheckBoxItem(sender as FrameworkElement);
        }

        /// <summary>
        /// Handles the Unchecked event of a tool window checkbox to hide the corresponding tool window.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void WindowCheckbox_Unchecked(object sender, RoutedEventArgs e)
        {
            DeselectCheckBoxItem(sender as FrameworkElement);
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
        /// Handles the Show All button click to display all available tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ShowAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteShowAllAvailableToolWindows();
        }

        /// <summary>
        /// Handles the Hide All button click to hide all available tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HideAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteHideAllAvailableToolWindows();
        }

        /// <summary>
        /// Handles the Stash button click to save the currently selected tool windows to a new stash.
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
        private void PopAbsButton_Click(object sender, RoutedEventArgs e)
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

            if (!(StashListBox.SelectedItem is ToolWindowStash stash))
            {
                return;
            }
            
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                ExecuteRestoreToolWindowsFromStash(stash);
            }
            else
            if (e.RightButton == MouseButtonState.Pressed)
            {
                ExecuteRestoreToolWindowsFromStashAbsolute(stash);
            }
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
        /// Executes the refresh operation and displays a status bar notification.
        /// </summary>
        private void ExecuteRefreshToolWindows()
        {
            RefreshToolWindows();

            StatusBarHelper.ShowStatusBarNotification("Tool windows refreshed.");
        }

        /// <summary>
        /// Shows all available tool windows and displays a status bar notification.
        /// </summary>
        private void ExecuteShowAllAvailableToolWindows()
        {
            SetAllToolWindowsVisibility(true);

            StatusBarHelper.ShowStatusBarNotification("All available tool windows shown.");
        }

        /// <summary>
        /// Hides all available tool windows and displays a status bar notification.
        /// </summary>
        private void ExecuteHideAllAvailableToolWindows()
        {
            SetAllToolWindowsVisibility(false);

            StatusBarHelper.ShowStatusBarNotification("All available tool windows hidden.");
        }

        /// <summary>
        /// Creates a new stash from the currently selected tool windows and displays a status bar notification.
        /// </summary>
        private void ExecuteStashSelectedToolWindows()
        {
            if (!IsInitialised)
            {
                return;
            }

            int selectedCount = ToolWindowsListBox.SelectedItems.Count;

            if (selectedCount == 0)
            {
                StatusBarHelper.ShowStatusBarNotification("No tool windows selected to stash. Select items in the list first.");
                return;
            }

            StashSelectedToolWindows();

            StatusBarHelper.ShowStatusBarNotification($"{selectedCount} selected tool window(s) stashed.");
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
        /// Clears all stashes after user confirmation and displays a status bar notification.
        /// </summary>
        private void ExecuteDropAllStashes()
        {
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
        /// Shows a tool window and adds it to the list box selection.
        /// </summary>
        /// <param name="sender">The framework element that triggered the event.</param>
        /// <exception cref="InvalidOperationException">Thrown when the sender's DataContext is not a ToolWindowEntry.</exception>
        private void SelectCheckBoxItem(FrameworkElement sender)
        {
            if (!(sender.DataContext is ToolWindowEntry entry))
            {
                throw new InvalidOperationException();
            }

            _toolWindowHelper.SetToolWindowVisibility(entry, true);

            if (!ToolWindowsListBox.SelectedItems.Contains(entry))
            {
                ToolWindowsListBox.SelectedItems.Add(entry);
            }
        }

        /// <summary>
        /// Hides a tool window and removes it from the list box selection.
        /// </summary>
        /// <param name="sender">The framework element that triggered the event.</param>
        private void DeselectCheckBoxItem(FrameworkElement sender)
        {
            if (!(sender.DataContext is ToolWindowEntry entry))
            {
                return;
            }

            _toolWindowHelper.SetToolWindowVisibility(entry, false);

            if (ToolWindowsListBox.SelectedItems.Contains(entry))
            {
                ToolWindowsListBox.SelectedItems.Remove(entry);
            }
        }

        /// <summary>
        /// Sets the visibility of all tool windows in the collection.
        /// </summary>
        /// <param name="isVisible">True to show all windows; false to hide all windows.</param>
        private void SetAllToolWindowsVisibility(bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            _toolWindowHelper.SetToolWindowsVisibility(ToolWindows, isVisible);
        }

        /// <summary>
        /// Creates a new stash from the currently selected tool windows.
        /// The stash is added to the top of the stash collection and persisted to settings.
        /// </summary>
        private void StashSelectedToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            if (ToolWindowsListBox.SelectedItems.Count == 0)
            {
                return;
            }

            var windowsToStash = ToolWindowsListBox.SelectedItems
                .Cast<ToolWindowEntry>()
                .OrderBy(entry => entry.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

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

        /// <summary>
        /// Clears all stashes after prompting the user for confirmation.
        /// </summary>
        private void DropAllStashes()
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

            Guid clsid = Guid.Empty;
            _uiShell.ShowMessageBox(
                0,
                ref clsid,
                "Drop All Stashes",
                "Are you sure you wish to drop all stashes?",
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
        /// Refreshes the tool windows list by querying Visual Studio for all available tool windows.
        /// Filters out unsupported windows and updates the UI with currently visible windows selected.
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

                if (entry.IsVisible)
                {
                    ToolWindowsListBox.SelectedItems.Add(entry);
                }
            }
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
            if (string.Equals(objectKindNormalized, ToggleToolWindowsToolWindow.ToolWindowGuidString, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (string.Equals(windowEntry.Caption, ToggleToolWindowsToolWindow.ToolWindowTitle, StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }
    #pragma warning restore VSTHRD010
    }
}