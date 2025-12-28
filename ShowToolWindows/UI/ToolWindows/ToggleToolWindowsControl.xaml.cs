using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
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

            RefreshToolWindows();
        }

#pragma warning disable VSTHRD010

        /// <summary>
        /// Handles the Checked event to show a tool window.
        /// </summary>
        private void WindowCheckbox_Checked(object sender, RoutedEventArgs e)
        {
            SetWindowVisibility(sender, true);
        }

        /// <summary>
        /// Handles the Unchecked event to hide a tool window.
        /// </summary>
        private void WindowCheckbox_Unchecked(object sender, RoutedEventArgs e)
        {
            SetWindowVisibility(sender, false);
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
            SetAllWindowsVisibility(true);
        }

        /// <summary>
        /// Handles the Hide All button click to hide all tool windows.
        /// </summary>
        private void HideAllButton_Click(object sender, RoutedEventArgs e)
        {
            SetAllWindowsVisibility(false);
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

            var windows = _dte.Windows
                .Cast<Window>()
                .Where(IsSupportedToolWindow)
                .OrderBy(w => w.Caption, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            foreach (Window window in windows)
            {
                ToolWindows.Add(new ToolWindowEntry(window));
            }
        }

        private void SelectAllToolWindows()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (ToolWindowEntry entry in ToolWindows)
            {
                entry.IsVisible = true;
                entry.SetVisibility(true);
                entry.Synchronize();
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

            if (string.Equals(window.ObjectKind, Constants.vsWindowKindMainWindow, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Exclude our own tool window by GUID
            string objectKindNormalized = window.ObjectKind.Trim('{', '}');
            if (string.Equals(objectKindNormalized, ToggleToolWindowsToolWindow.ToolWindowGuidString,
                    StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            // Also exclude by caption as a fallback
            if (string.Equals(window.Caption, ToggleToolWindowsToolWindow.ToolWindowTitle, StringComparison.Ordinal))
            {
                return false;
            }

            return true;
        }

        private void SetWindowVisibility(object sender, bool isVisible)
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

        private void SetAllWindowsVisibility(bool isVisible)
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