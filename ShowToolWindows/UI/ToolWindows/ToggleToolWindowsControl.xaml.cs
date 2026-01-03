using EnvDTE;
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
            _dte = dteService as DTE;
            _uiShell = await _package.GetServiceAsync(typeof(SVsUIShell)) as IVsUIShell;

            _stashService = new StashSettingsService(_package);
            _toolWindowHelper = new ToolWindowHelper(_dte, _uiShell);
            LoadToolWindowStashes();

            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        private void ToggleToolWindowsControl_Loaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged += Stashes_CollectionChanged;
        }

        private void ToggleToolWindowsControl_Unloaded(object sender, RoutedEventArgs e)
        {
            Stashes.CollectionChanged -= Stashes_CollectionChanged;
        }

        private void Stashes_CollectionChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            OnPropertyChanged(nameof(StashesHeader));
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

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

        private void StashOpenToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var visibleWindows = _dte.Windows
                .Cast<Window>()
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
                CreatedAt = DateTimeOffset.UtcNow
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

            for (int i = 0; i < objectKinds.Length; i++)
            {
                var objectKind = objectKinds[i];
                var caption = i < captions.Length ? captions[i] : "Unknown";

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

#pragma warning restore VSTHRD010
    }
}