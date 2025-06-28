using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows
{
    /// <summary>
    /// Command handler
    /// </summary>

    #pragma warning disable VSTHRD010 // Invoke single-threaded types on Main thread
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
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowSolutionExplorerCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ShowSolutionExplorerCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
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
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
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
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            Debug.WriteLine("ShowSolutionExplorerCommand.Execute called");

            try
            {
                // Get the Solution Explorer tool window
                var dte = Package.GetGlobalService(typeof(EnvDTE.DTE)) as EnvDTE.DTE;
                if (dte == null)
                {
                    Debug.WriteLine("ERROR: Could not get DTE service");
                    return;
                }

                Debug.WriteLine("Got DTE service successfully");

                // Get the IDE window bounds
                var visualStudioWindow = dte.MainWindow;

                // Show the Solution Explorer window
                var solutionWindow = dte.Windows.Item(EnvDTE.Constants.vsWindowKindSolutionExplorer);
                if (solutionWindow == null)
                {
                    Debug.WriteLine("ERROR: Solution Explorer window not found");
                    return;
                }

                Debug.WriteLine("Found Solution Explorer window");
                solutionWindow.Visible = true;
                solutionWindow.Activate();

                RepositionIfOffscreen(visualStudioWindow, solutionWindow, SystemInformation.VirtualScreen);
            }
            catch (Exception ex)
            {
                // Log the error or show a message
                Debug.WriteLine($"ERROR showing Solution Explorer: {ex.Message}");
                Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        private void RepositionIfOffscreen(Window visualStudioWindow, Window windowToRepos, Rectangle screen)
        {

            int winLeft = windowToRepos.Left;
            int winTop = windowToRepos.Top;
            int winWidth = windowToRepos.Width;
            int winHeight = windowToRepos.Height;
            int winRight = winLeft + winWidth;
            int winBottom = winTop + winHeight;

            // If Visual Studio's bottom edge is off screen, position solution explorer at the top of the screen
            // Otherwise, position at Visual Studio's top
            var y = (visualStudioWindow.Top + visualStudioWindow.Height) > screen.Bottom ? screen.Top : visualStudioWindow.Top; 

            // Any part of the solution explorer off left edge of screen?
            if (winLeft < screen.Left)
            {
                // Float and move to left edge
                FloatWindowAt(windowToRepos, screen.Left, y, winWidth, winHeight, screen);
                Debug.WriteLine("Solution Explorer floated and positioned at left edge of screen");
                return;
            }

            // Any part of the solution explorer off right edge of screen?
            if (winRight > screen.Right)
            {
                // Float and move to right edge
                FloatWindowAt(windowToRepos, screen.Right - winWidth, y, winWidth, winHeight, screen);
                Debug.WriteLine("Solution Explorer floated and positioned at right edge of screen");
                return;
            }

            // Any part of the solution explorer off bottom edge of screen?  
            if (winBottom > screen.Bottom)
            {
                // Float and move to top edge
                FloatWindowAt(windowToRepos, winLeft, y, winWidth, winHeight, screen);
                Debug.WriteLine("Solution Explorer floated and positioned at top of screen");
                return;
            }
        }

        // Helper to float and position the window
        private void FloatWindowAt(Window window, int left, int top, int width, int height, Rectangle screen)
        {
            window.IsFloating = true;
            // Clamp width/height to screen
            int screenHeight = screen.Height;
            window.Width = Math.Min(width, 400); // reasonable default width
            window.Height = Math.Min(height, screenHeight);
            window.Left = left;
            window.Top = top;
            window.Activate();
        }
    }
}
