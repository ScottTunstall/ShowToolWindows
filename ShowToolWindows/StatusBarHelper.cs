using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Diagnostics;

namespace ShowToolWindows
{
    internal static class StatusBarHelper
    {
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
