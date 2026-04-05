using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using ShowToolWindows.Model;
using ShowToolWindows.Services;
using ShowToolWindows.UI.Infrastructure;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using Task = System.Threading.Tasks.Task;

namespace ShowToolWindows.Commands
{
    /// <summary>
    /// Command handler that populates the Apply Stash submenu with dynamic items
    /// representing the most recent stashes and applies the selected stash in absolute mode.
    /// </summary>
    internal sealed class ApplyStashSubmenuCommand
    {
        private const int MaxStashItems = 10;
        private const int MaxMenuItemTextLength = 80;

        /// <summary>
        /// Base command ID for the dynamic stash menu items.
        /// </summary>
        private const int BaseCommandId = 0x2000;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        private static readonly Guid _commandSet = new Guid("a2a7f2af-5e8b-4a37-a078-5a8dbe625606");

        private readonly AsyncPackage _package;
        private StashSettingsService _stashService;

        private ApplyStashSubmenuCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            _package = package ?? throw new ArgumentNullException(nameof(package));
            if (commandService == null)
            {
                throw new ArgumentNullException(nameof(commandService));
            }

            var dynamicItemRootId = new CommandID(_commandSet, BaseCommandId);
            var dynamicCommand = new DynamicItemMenuCommand(
                dynamicItemRootId,
                OnInvokedDynamicItem,
                OnBeforeQueryStatusDynamicItem,
                cmdId => cmdId >= BaseCommandId && cmdId < BaseCommandId + MaxStashItems);

            commandService.AddCommand(dynamicCommand);
        }

        /// <summary>
        /// Gets the singleton instance of this command.
        /// </summary>
        public static ApplyStashSubmenuCommand Instance { get; private set; }

        /// <summary>
        /// Initialises the command and registers it with the Visual Studio command service.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;

            Instance = new ApplyStashSubmenuCommand(package, commandService);
        }

        /// <summary>
        /// Event handler for dynamic item invocation. Delegates to the execute method.
        /// </summary>
        private void OnInvokedDynamicItem(object sender, EventArgs e)
        {
            ExecuteApplyStash(sender);
        }

        /// <summary>
        /// Event handler for query status. Delegates to the execute method.
        /// </summary>
        private void OnBeforeQueryStatusDynamicItem(object sender, EventArgs e)
        {
            ExecuteUpdateDynamicMenuItemStatus(sender);
        }

        /// <summary>
        /// Sets the text, visibility and enabled state for each dynamic menu item
        /// based on the current stash collection.
        /// </summary>
        private void ExecuteUpdateDynamicMenuItemStatus(object sender)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var command = (DynamicItemMenuCommand)sender;
            int index = GetStashIndex(command);

            var stashes = GetStashService().LoadStashes();
            int itemCount = Math.Min(stashes.Count, MaxStashItems);

            if (index >= 0 && index < itemCount)
            {
                command.Visible = true;
                command.Enabled = true;
                var text = $"{{{index}}} - {stashes[index]}";
                command.Text = text.Length > MaxMenuItemTextLength
                    ? text.Substring(0, MaxMenuItemTextLength - 3) + "..."
                    : text;
            }
            else
            {
                command.Visible = false;
                command.Enabled = false;
                command.MatchedCommandId = 0;
            }
        }

        /// <summary>
        /// Returns the <see cref="StashSettingsService"/>, creating it on first use.
        /// </summary>
        private StashSettingsService GetStashService()
        {
            return _stashService ?? (_stashService = new StashSettingsService(_package));
        }

        /// <summary>
        /// Applies the selected stash in absolute mode and shows a status bar notification.
        /// </summary>
        private void ExecuteApplyStash(object sender)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var command = (DynamicItemMenuCommand)sender;
            int index = GetStashIndex(command);

            var stashes = GetStashService().LoadStashes();
            if (index < 0 || index >= stashes.Count)
            {
                return;
            }

            var stash = stashes[index];
            ApplyStashAbsolute(stash);

            StatusBarHelper.ShowStatusBarNotification($"Tool windows replaced by stash@{{{index}}}.");
        }

        /// <summary>
        /// Closes all tool windows not present in the stash, then restores the stash windows.
        /// </summary>
        private void ApplyStashAbsolute(ToolWindowStash stash)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!(Package.GetGlobalService(typeof(DTE)) is DTE dte))
            {
                Debug.WriteLine("ERROR: Could not get DTE service");
                StatusBarHelper.ShowStatusBarNotification("Error: Could not access Visual Studio services");
                return;
            }

            if (!(Package.GetGlobalService(typeof(SVsUIShell)) is IVsUIShell uiShell))
            {
                Debug.WriteLine("ERROR: Could not get IVsUIShell service");
                StatusBarHelper.ShowStatusBarNotification("Error: Could not access Visual Studio services");
                return;
            }

            var helper = ToolWindowHelper.Create(dte, uiShell);
            helper.CloseToolWindowsNotInStash(stash);
            helper.RestoreToolWindowsFromStash(stash);
        }

        /// <summary>
        /// Calculates the stash index from the matched command identifier.
        /// </summary>
#pragma warning disable SA1204

        private static int GetStashIndex(DynamicItemMenuCommand command)
#pragma warning restore SA1204
        {
            return command.MatchedCommandId == 0
                ? 0
                : command.MatchedCommandId - BaseCommandId;
        }
    }
}