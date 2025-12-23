using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;

namespace ShowToolWindows.UI
{
    /// <summary>
    /// Provides helper methods for writing messages to the Visual Studio status bar.
    /// </summary>
    internal static class StatusBarHelper
    {
        /// <summary>
        /// Shows a notification message in the Visual Studio status bar.
        /// </summary>
        /// <param name="message">The text to display in the status bar.</param>
        /// <remarks>
        /// This method must be called on the UI thread. If the status bar service is unavailable
        /// or an error occurs, the error will be written to the debug output without throwing an exception.
        /// </remarks>
        public static void ShowStatusBarNotification(string message)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (Package.GetGlobalService(typeof(SVsStatusbar)) is IVsStatusbar statusBar)
                {
                    statusBar.SetText(message);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error showing notification: {ex.Message}");
            }
        }
    }
}
