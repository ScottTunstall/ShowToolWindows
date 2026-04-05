using EnvDTE;
using Microsoft.VisualStudio.Shell;
using ShowToolWindows.UI.Infrastructure;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// Command handler that closes all visible tool windows in Visual Studio except Solution Explorer.
    /// </summary>
    /// <remarks>
    /// This command closes all tool windows except the main window and Solution Explorer, providing a way to
    /// quickly clean up the Visual Studio workspace while preserving access to the solution structure.
    /// </remarks>
    internal sealed class CloseAllToolWindowsExceptSolutionExplorerCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4129;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a2a7f2af-5e8b-4a37-a078-5a8dbe625606");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseAllToolWindowsExceptSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CloseAllToolWindowsExceptSolutionExplorerCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static CloseAllToolWindowsExceptSolutionExplorerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CloseAllButSolutionExplorerCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CloseAllToolWindowsExceptSolutionExplorerCommand(package, commandService);
        }

        /// <summary>
        /// Handles the menu item click to close all tool windows except Solution Explorer.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ExecuteCloseAllToolWindowsExceptSolutionExplorer();
        }

        /// <summary>
        /// Closes all visible tool windows except Solution Explorer and displays a status bar notification.
        /// </summary>
        private void ExecuteCloseAllToolWindowsExceptSolutionExplorer()
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

                var excluded = new HashSet<string> { Constants.vsWindowKindSolutionExplorer };
                int closedCount = WindowHelper.CloseVisibleToolWindows(dte, excluded);

                string message = closedCount == 1
                    ? "Closed 1 tool window."
                    : $"Closed {closedCount} tool windows.";
                StatusBarHelper.ShowStatusBarNotification(message);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ERROR closing tool windows: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                StatusBarHelper.ShowStatusBarNotification($"Error closing tool windows: {ex.Message}");
            }
        }
    }
}