using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Provides helper methods for displaying Visual Studio message boxes.
    /// </summary>
    internal class MessageBoxHelper
    {
        private readonly IVsUIShell _uiShell;

        /// <summary>
        /// Initializes a new instance of the <see cref="MessageBoxHelper"/> class.
        /// </summary>
        /// <param name="uiShell">The Visual Studio UI shell service.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="uiShell"/> is null.</exception>
        public MessageBoxHelper(IVsUIShell uiShell)
        {
            _uiShell = uiShell ?? throw new ArgumentNullException(nameof(uiShell));
        }

        /// <summary>
        /// Shows a confirmation dialog with Yes, No, and Cancel options.
        /// </summary>
        /// <param name="title">The dialog title.</param>
        /// <param name="message">The dialog message body.</param>
        /// <returns>The selected <see cref="VSConstants.MessageBoxResult"/>.</returns>
        public VSConstants.MessageBoxResult ShowConfirmModalDialog(string title, string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Guid clsid = Guid.Empty;
            _uiShell.ShowMessageBox(
                0,
                ref clsid,
                title,
                message,
                null,
                0,
                OLEMSGBUTTON.OLEMSGBUTTON_YESNOCANCEL,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_SECOND,
                OLEMSGICON.OLEMSGICON_QUERY,
                0,
                out int result);

            return (VSConstants.MessageBoxResult)result;
        }
    }
}
