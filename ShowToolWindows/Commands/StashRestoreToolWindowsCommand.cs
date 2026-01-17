using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.UI.ToolWindows;
using System;
using System.ComponentModel.Design;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// Command handler that shows the Stash/Restore tool windows tool window.
    /// </summary>
    internal sealed class StashRestoreToolWindowsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 4131;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("a2a7f2af-5e8b-4a37-a078-5a8dbe625606");

        private readonly AsyncPackage _package;

        private StashRestoreToolWindowsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            if (commandService == null)
            {
                throw new ArgumentNullException(nameof(commandService));
            }

            var menuCommandId = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand((s, e) => ThreadHelper.JoinableTaskFactory.Run(async() => await ExecuteAsync()), menuCommandId);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the singleton instance of the command.
        /// </summary>
        public static StashRestoreToolWindowsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Initializes the command and registers it with the Visual Studio command service.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new StashRestoreToolWindowsCommand(package, commandService);
        }

        private async Task ExecuteAsync()
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(_package.DisposalToken);

            ToolWindowPane window = await _package.ShowToolWindowAsync(typeof(StashRestoreToolWindowsToolWindow), 0, true, _package.DisposalToken);
            if (window?.Frame == null)
            {
                throw new NotSupportedException("Cannot create Stash/Restore tool windows tool window.");
            }

            StashRestoreToolWindowsToolWindow pane = (StashRestoreToolWindowsToolWindow)window;
            await pane.InitializeAsync(_package);

            IVsWindowFrame windowFrame = (IVsWindowFrame)window.Frame;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(windowFrame.Show());
        }
    }
}
