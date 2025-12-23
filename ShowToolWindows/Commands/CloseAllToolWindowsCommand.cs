using EnvDTE;
using Microsoft.VisualStudio.Shell;
using ShowToolWindows.UI;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Linq;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// Command handler that closes all visible tool windows in Visual Studio.
    /// </summary>
    /// <remarks>
    /// This command closes all tool windows except the main window, providing a way to quickly 
    /// clean up the Visual Studio workspace by hiding all visible tool windows at once.
    /// </remarks>
    internal sealed class CloseAllToolWindowsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4130;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a2a7f2af-5e8b-4a37-a078-5a8dbe625606");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseAllToolWindowsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CloseAllToolWindowsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(Execute, menuCommandId);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CloseAllToolWindowsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return _package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CloseAllToolWindowsCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CloseAllToolWindowsCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                {
                    Debug.WriteLine("ERROR: Could not get DTE service");
                    StatusBarHelper.ShowStatusBarNotification("Error: Could not access Visual Studio services");
                    return;
                }

#pragma warning disable VSTHRD010
                var windowsToClose = dte.Windows
                    .Cast<Window>()
                    .Where(w=> w.Visible)
                    .Where(w => w.Kind == WindowKindConsts.ToolWindowKind)
                    .Where(w=> w.ObjectKind != Constants.vsWindowKindMainWindow)
                    .ToList();
#pragma warning restore VSTHRD010

                int closedCount = WindowHelper.CloseWindows(windowsToClose);
                
                string message = closedCount == 1 
                    ? "Closed 1 tool window." 
                    : $"Closed {closedCount} tool windows.";
                StatusBarHelper.ShowStatusBarNotification(message);
            }
            catch (Exception ex)
            {
                // Log the error or show a message
                Debug.WriteLine($"ERROR closing tool windows: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                StatusBarHelper.ShowStatusBarNotification($"Error closing tool windows: {ex.Message}");
            }
        }
    }
}
