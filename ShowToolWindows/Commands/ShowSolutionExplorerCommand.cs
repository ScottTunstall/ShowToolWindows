using EnvDTE;
using Microsoft.VisualStudio.Shell;
using ShowToolWindows.UI.Infrastructure;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// Command handler that shows and activates the Solution Explorer tool window.
    /// </summary>
    /// <remarks>
    /// This command makes the Solution Explorer visible and brings it to the foreground.
    /// If the Solution Explorer is positioned off-screen, the command automatically repositions
    /// it to be visible within the current screen bounds.
    /// </remarks>
    internal sealed class ShowSolutionExplorerCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a2a7f2af-5e8b-4a37-a078-5a8dbe625606");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage _package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ShowSolutionExplorerCommand(AsyncPackage package, OleMenuCommandService commandService)
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
        public static ShowSolutionExplorerCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initialises the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ShowSolutionExplorerCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ShowSolutionExplorerCommand(package, commandService);
        }

        /// <summary>
        /// Handles the menu item click to show the Solution Explorer window.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ExecuteShowSolutionExplorer();
        }

        /// <summary>
        /// Shows the Solution Explorer window, repositioning it if off-screen.
        /// </summary>
        private void ExecuteShowSolutionExplorer()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
                {
                    Debug.WriteLine("ERROR: Could not get DTE service");
                    return;
                }

                var visualStudioWindow = dte.MainWindow;

                var solutionWindow = dte.Windows.Item(Constants.vsWindowKindSolutionExplorer);
                if (solutionWindow == null)
                {
                    Debug.WriteLine("ERROR: Solution Explorer window not found");
                    return;
                }

                solutionWindow.Visible = true;
                solutionWindow.Activate();

                WindowHelper.RepositionIfOffscreen(visualStudioWindow, solutionWindow, SystemInformation.VirtualScreen);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"ERROR showing Solution Explorer: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }
    }
}