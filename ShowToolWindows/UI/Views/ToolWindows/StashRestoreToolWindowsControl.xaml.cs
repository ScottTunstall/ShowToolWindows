using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using ShowToolWindows.Services;
using ShowToolWindows.UI.Infrastructure;
using ShowToolWindows.UI.ViewModels;
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

namespace ShowToolWindows.UI.Views.ToolWindows
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
        private MessageBoxHelper _messageBoxHelper;
        private bool _isInitialised;
        private bool _hasSelectedItems;
        private readonly string _toolWindowObjectKind;


        /// <summary>
        /// Initializes a new instance of the <see cref="ToggleToolWindowsControl"/> class.
        /// </summary>
        public ToggleToolWindowsControl()
        {
            InitializeComponent();

            _toolWindowObjectKind = ObjectKindHelper.NormalizeObjectKind(StashRestoreToolWindowsToolWindow.ToolWindowGuidString);

            DataContext = this;
            RefreshCommand = new RelayCommand(parameter =>
            {
                if (!IsInitialised)
                {
                    return;
                }

                ExecuteRefresh();
            });
            DropStashCommand = new RelayCommand(parameter =>
            {
                if (!IsInitialised)
                {
                    return;
                }

                ExecuteDropStash();
            });
            CheckAllCommand = new RelayCommand(parameter =>
            {
                if (!IsInitialised)
                {
                    return;
                }

                ExecuteCheckAll();
            });

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
        public ToolWindowStashCollection Stashes
        {
            get;
        } = new ToolWindowStashCollection();

        /// <summary>
        /// Gets the collection of stash display items with dynamic indices for UI binding.
        /// </summary>
        // ReSharper disable once CollectionNeverQueried.Global
        // ReSharper disable once MemberCanBePrivate.Global
        public ObservableCollection<StashListItem> StashListItems
        {
            get;
        } = new ObservableCollection<StashListItem>();

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
            _messageBoxHelper = new MessageBoxHelper(_uiShell);
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
            ThreadHelper.ThrowIfNotOnUIThread();

            OnPropertyChanged(nameof(StashesHeader));
            OnPropertyChanged(nameof(HaveStashes));
            OnPropertyChanged(nameof(CanMutateStash));
            RebuildStashListItems();
        }

        /// <summary>
        /// Rebuilds the StashListItems collection and updates indices dynamically.
        /// </summary>
        private void RebuildStashListItems()
        {
            StashListItems.Clear();
            for (int i = 0; i < Stashes.Count; i++)
            {
                StashListItems.Add(new StashListItem(Stashes[i], i));
            }
        }

        /// <summary>
        /// Handles the CollectionChanged event of the ToolWindows collection to subscribe to item property changes.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="NotifyCollectionChangedEventArgs"/> instance containing the event data.</param>
        private void ToolWindows_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
            ThreadHelper.ThrowIfNotOnUIThread();

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
            if (!IsInitialised)
            {
                return;
            }

            ExecuteRefreshToolWindows();
        }

        /// <summary>
        /// Handles the Stash button click to save the currently checked tool windows to a new stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void StashButton_Click(object sender, RoutedEventArgs e)
        {
            if (!IsInitialised)
            {
                return;
            }

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
            if (!IsInitialised)
            {
                return;
            }

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
            if (!IsInitialised)
            {
                return;
            }

            ExecutePopToolWindowsFromStashAbsolute();
        }

        /// <summary>
        /// Handles the Drop All button click to clear all stashes after user confirmation.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DropAllButton_Click(object sender, RoutedEventArgs e)
        {
            if (!IsInitialised)
            {
                return;
            }

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
            if (!IsInitialised)
            {
                return;
            }

            if (e.LeftButton == MouseButtonState.Pressed)
            {
                ExecuteRestoreToolWindowsFromStash();
            }
        }

        /// <summary>
        /// Handles right-click on a stash item to select it before the context menu appears.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="MouseButtonEventArgs"/> instance containing the event data.</param>
        private void StashListBox_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
            if (!IsInitialised)
            {
                return;
            }

            ExecuteRestoreToolWindowsFromStashAbsolute();
        }

        /// <summary>
        /// Handles the Apply (Merge) context menu click to restore tool windows from the selected stash in merge mode.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ApplyStashMerge_Click(object sender, RoutedEventArgs e)
        {
            if (!IsInitialised)
            {
                return;
            }

            ExecuteRestoreToolWindowsFromStash();
        }

        /// <summary>
        /// Handles the Hide all visible context menu click to hide all tool windows referenced in the selected stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HideAllVisible_Click(object sender, RoutedEventArgs e)
        {
            if (!IsInitialised)
            {
                return;
            }

            ExecuteHideAllVisibleInStash();
        }

        /// <summary>
        /// Handles the Drop context menu click to remove the selected stash.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DropStash_Click(object sender, RoutedEventArgs e)
        {
            if (!IsInitialised)
            {
                return;
            }

            ExecuteDropStash();
        }

        /// <summary>
        /// Handles digit key presses to select a stash by index from anywhere in the control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The key event data.</param>
        private void UserControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!IsInitialised || Keyboard.Modifiers != ModifierKeys.None)
            {
                return;
            }

            if (!KeyHelper.TryGetDigitFromKey(e.Key, out int index))
            {
                return;
            }

            ExecuteSelectStashByIndex(index);
            e.Handled = true;
        }


        /// <summary>
        /// Executes the refresh command to reload the tool windows list.
        /// </summary>
        private void ExecuteRefresh()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            RefreshToolWindows();
        }

        /// <summary>
        /// Executes the check all command to select all tool windows.
        /// </summary>
        private void ExecuteCheckAll()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
            ThreadHelper.ThrowIfNotOnUIThread();

            RefreshToolWindows();

            StatusBarHelper.ShowStatusBarNotification("Tool windows refreshed.");
        }

        /// <summary>
        /// Creates a new stash from the currently checked tool windows and displays a status bar notification.
        /// </summary>
        private void ExecuteStashSelectedToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
        /// Selects the stash list item at the specified index and applies it in merge mode, if present.
        /// </summary>
        /// <param name="index">The stash index to select.</param>
        private void ExecuteSelectStashByIndex(int index)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (index < 0 || index >= StashListItems.Count)
            {
                return;
            }

            StashListBox.SelectedIndex = index;
            StashListBox.ScrollIntoView(StashListBox.SelectedItem);

            if (!(StashListBox.SelectedItem is StashListItem listItem))
            {
                return;
            }

            ExecuteRestoreToolWindowsFromStash();
            StatusBarHelper.ShowStatusBarNotification($"Tool windows merged from stash@{index}.");
        }

        /// <summary>
        /// Restores tool windows from the top stash in merge mode (keeps existing windows open) and displays a status bar notification.
        /// </summary>
        private void ExecutePopAndMergeToolWindowsFromStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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
            ThreadHelper.ThrowIfNotOnUIThread();

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

            if (!(this.StashListBox.SelectedItem is StashListItem listItem))
            {
                return;
            }

            var stash = listItem.Stash;
            var result = _messageBoxHelper.ShowConfirmModalDialog(
                "Drop Stash",
                "Are you sure you wish to drop this stash?");

            if (result != VSConstants.MessageBoxResult.IDYES)
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

            var result = _messageBoxHelper.ShowConfirmModalDialog(
                "Drop All Stashes",
                "Are you sure you wish to drop all stashes? This cannot be undone.");

            if (result != VSConstants.MessageBoxResult.IDYES)
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
        private void ExecuteRestoreToolWindowsFromStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is StashListItem listItem))
            {
                return;
            }

            var stash = listItem.Stash;
            RestoreToolWindowsFromStash(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows restored from stash.");
        }

        /// <summary>
        /// Restores tool windows from a specific stash in absolute mode (closes windows not in the stash).
        /// Does not remove the stash.
        /// </summary>
        private void ExecuteRestoreToolWindowsFromStashAbsolute()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is StashListItem listItem))
            {
                return;
            }

            var stash = listItem.Stash;

            RestoreToolWindowsFromStashAbsolute(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows restored from stash.");
        }

        /// <summary>
        /// Hides all visible tool windows that are referenced in the specified stash.
        /// </summary>
        private void ExecuteHideAllVisibleInStash()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(this.StashListBox.SelectedItem is StashListItem listItem))
            {
                return;
            }

            var stash = listItem.Stash;

            HideAllVisibleInStash(stash);
            StatusBarHelper.ShowStatusBarNotification("Tool windows in stash hidden.");
        }

        /// <summary>
        /// Creates a new stash from the currently checked tool windows.
        /// The stash is added to the top of the stash collection and persisted to settings.
        /// </summary>
        private void StashSelectedToolWindows()
        {
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

            if (Stashes.HasMatchingConfiguration(objectKinds) && !ConfirmCreateDuplicateStash(captions))
            {
                return;
            }

            Stashes.CreateAndPushToTop(captions, objectKinds);

            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Prompts the user before creating a duplicate stash configuration.
        /// </summary>
        /// <param name="captions">The tool window captions to display in the prompt.</param>
        /// <returns>True if the user selects Yes; otherwise, false.</returns>
        private bool ConfirmCreateDuplicateStash(IReadOnlyList<string> captions)
        {
            string captionList = string.Join(Environment.NewLine, captions.Select(c => "- " + c));
            string message = "A stash with the following tool windows already exists. Push another?"
                             + Environment.NewLine
                             + Environment.NewLine
                             + captionList;

            var result = _messageBoxHelper.ShowConfirmModalDialog("Duplicate Stash", message);

            return result == VSConstants.MessageBoxResult.IDYES;
        }


        /// <summary>
        /// Restores tool windows from the top stash in merge mode and removes the stash.
        /// Does not close windows that are not in the stash.
        /// </summary>
        private void PopToolWindowsFromStash()
        {
            var stash = Stashes[0];
            RestoreToolWindowsFromStash(stash);
            Stashes.DeleteTopOfStack();
            SaveAllToolWindowStashes();
        }

        /// <summary>
        /// Restores tool windows from the top stash in absolute mode and removes the stash.
        /// Closes all windows not in the stash before restoring.
        /// </summary>
        private void PopAbsToolWindowsFromStash()
        {
            var stash = Stashes[0];
            RestoreToolWindowsFromStashAbsolute(stash);
            Stashes.DeleteTopOfStack();
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
            _toolWindowHelper.RestoreToolWindowsFromStash(stash);

            // Make the tool window active again after restoring from stash
            _toolWindowHelper.TryOpenToolWindowByObjectKind(_toolWindowObjectKind);

            RefreshToolWindows();
        }

        /// <summary>
        /// Restores tool windows from the given stash in absolute mode and refreshes the tool windows list.
        /// Closes all windows not in the stash before restoring.
        /// </summary>
        /// <param name="stash">The stash containing the tool windows to restore.</param>
        private void RestoreToolWindowsFromStashAbsolute(ToolWindowStash stash)
        {
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