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
using Window = EnvDTE.Window;

namespace ShowToolWindows.UI.ToolWindows
{
    /// <summary>
    /// Interaction logic for CloseToolWindowsControl.
    /// </summary>
    public partial class ToggleToolWindowsControl : UserControl, INotifyPropertyChanged
    {
        private AsyncPackage _package;
        private DTE _dte;
        private IVsUIShell _uiShell;
        private StashSettingsService _stashService;
        private ToolWindowHelper _toolWindowHelper;
        private bool _isInitialised;

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
        /// Used by the F5 keybinding to refresh the tool windows list.
        /// Gets the command for refreshing the tool windows list.
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
        /// Gets the header text for the stashes expander, including the count.
        /// </summary>
        public string StashesHeader => $"Stashes ({Stashes.Count})";

        /// <summary>
        /// Gets the collection of stashed tool window snapshots.
        /// </summary>
        public ObservableCollection<ToolWindowStash> Stashes
        {
            get;
        } = new ObservableCollection<ToolWindowStash>();

        /// <summary>
        /// Gets a value indicating whether the control has been initialised with VS services.
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
                }
            }
        }

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;


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
            _dte = dteService as DTE ?? throw new InvalidOperationException("Failed to get DTE service.");
            
            _uiShell = await _package.GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell ?? throw new InvalidOperationException("Failed to get IVsUIShell service.");

            _stashService = new StashSettingsService(_package);
            _toolWindowHelper = new ToolWindowHelper(_dte, _uiShell);
            LoadToolWindowStashes();

            IsInitialised = true;
            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        /// <summary>
        /// Handles the Loaded event of the control to wire up event subscriptions.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ToggleToolWindowsControl_Loaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged += Stashes_CollectionChanged;
        }

        /// <summary>
        /// Handles the Unloaded event of the control to remove event subscriptions.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ToggleToolWindowsControl_Unloaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged -= Stashes_CollectionChanged;
        }

        /// <summary>
        /// Handles the CollectionChanged event of the Stashes collection to update the stashes header text.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="NotifyCollectionChangedEventArgs"/> instance containing the event data.</param>
        private void Stashes_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(StashesHeader));
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
        /// Handles the Checked event to show a tool window.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void WindowCheckbox_Checked(object sender, RoutedEventArgs e)
        {
            SetToolWindowVisibility(sender, true);
        }

        /// <summary>
        /// Handles the Unchecked event to hide a tool window.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void WindowCheckbox_Unchecked(object sender, RoutedEventArgs e)
        {
            SetToolWindowVisibility(sender, false);
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
        /// Handles the Show All button click to display all tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void ShowAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteShowAllAvailableToolWindows();
        }


        /// <summary>
        /// Handles the Hide All button click to hide all tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void HideAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteHideAllAvailableToolWindows();
        }

        /// <summary>
        /// Handles the Stash button click to save current visible tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void StashButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteStashVisibleToolWindows();
        }

        /// <summary>
        /// Handles the Pop button click to restore and remove the top stashed item.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void PopButton_Click(object sender, RoutedEventArgs e)
        {
            ExecutePopToolWindowsFromStash();
        }

        /// <summary>
        /// Handles the Drop All button click to clear all stashes after confirmation.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RoutedEventArgs"/> instance containing the event data.</param>
        private void DropAllButton_Click(object sender, RoutedEventArgs e)
        {
            ExecuteDropAllStashes();
        }


        /// <summary>
        /// Handles double-click on a stash to restore those tool windows.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="MouseButtonEventArgs"/> instance containing the event data.</param>
        private void StashListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (StashListBox.SelectedItem is ToolWindowStash stash)
            {
                RestoreToolWindowStash(stash);
            }
        }

        /// <summary>
        /// Executes the refresh command to reload the tool windows list.
        /// </summary>
        /// <param name="parameter">The command parameter.</param>
        private void ExecuteRefresh(object parameter)
        {
            RefreshToolWindows();
        }


        private void ExecuteRefreshToolWindows()
        {
            RefreshToolWindows();

            StatusBarHelper.ShowStatusBarNotification("Tool windows refreshed.");
        }

        private void ExecuteShowAllAvailableToolWindows()
        {
            SetAllToolWindowsVisibility(true);

            StatusBarHelper.ShowStatusBarNotification("All available tool windows shown.");
        }

        private void ExecuteHideAllAvailableToolWindows()
        {
            SetAllToolWindowsVisibility(false);

            StatusBarHelper.ShowStatusBarNotification("All available tool windows hidden.");
        }


        private void ExecuteStashVisibleToolWindows()
        {
            StashVisibleToolWindows();

            StatusBarHelper.ShowStatusBarNotification("Tool windows stashed.");
        }

        private void ExecutePopToolWindowsFromStash()
        {
            if (Stashes.Count == 0)
            {
                return;
            }

            PopToolWindowsFromStash();

            StatusBarHelper.ShowStatusBarNotification("Tool windows popped from stash.");
        }

        private void ExecuteDropAllStashes()
        {
            DropAllStashes();

            StatusBarHelper.ShowStatusBarNotification("All tool window stashes dropped.");
        }


        private void RefreshToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            ToolWindows.Clear();

            var allWindows = _dte.Windows
                .Cast<Window>()
                .ToList();
            
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

        private void StashVisibleToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            var selectedWindows = ToolWindows
                .Where(entry => entry.IsVisible)
                .OrderBy(entry => entry.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            if (selectedWindows.Count == 0)
            {
                return;
            }

            var captions = new List<string>();
            var objectKinds = new List<string>();

            foreach (var entry in selectedWindows)
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
            SaveToolWindowStashes();
        }

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
            RestoreToolWindowStash(stash);
            Stashes.RemoveAt(0);
            SaveToolWindowStashes();
        }

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
            SaveToolWindowStashes();
        }

        private void LoadToolWindowStashes()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

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

            if (!IsInitialised)
            {
                return;
            }

            var stashList = new List<ToolWindowStash>(Stashes);
            _stashService.SaveStashes(stashList);
        }

        private void RestoreToolWindowStash(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            var objectKinds = stash.WindowObjectKinds;

            for (int i = 0; i < objectKinds.Length; i++)
            {
                var objectKind = objectKinds[i];

                _toolWindowHelper.TryOpenToolWindowByObjectKind(objectKind);
            }

            RefreshToolWindows();
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
            if (string.Equals(objectKindNormalized, ToggleToolWindowsToolWindow.ToolWindowGuidString, StringComparison.OrdinalIgnoreCase))
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

            if (!IsInitialised)
            {
                return;
            }

            FrameworkElement frameworkElement = sender as FrameworkElement;
            if (!(frameworkElement?.DataContext is ToolWindowEntry entry))
            {
                return;
            }

            _toolWindowHelper.SetToolWindowVisibility(entry, isVisible);
        }

        private void SetAllToolWindowsVisibility(bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!IsInitialised)
            {
                return;
            }

            _toolWindowHelper.SetToolWindowsVisibility(ToolWindows, isVisible);
        }


#pragma warning restore VSTHRD010
    }
}