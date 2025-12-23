using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace ShowToolWindows
{
    /// <summary>
    /// Helper class for managing Visual Studio tool windows.
    /// </summary>
    internal static class WindowHelper
    {
        /// <summary>
        /// Closes the specified collection of Visual Studio tool windows.
        /// </summary>
        /// <param name="windowsToClose">The collection of windows to close.</param>
        /// <returns>The number of windows that were successfully closed.</returns>
        /// <remarks>
        /// Only windows that are currently visible will be closed and counted in the return value.
        /// Windows that fail to close will be logged to the debug output but will not be counted.
        /// </remarks>
        public static int CloseWindows(IEnumerable<EnvDTE.Window> windowsToClose)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var asList = windowsToClose.ToList();

            int closedCount = 0;
            for (int i = asList.Count - 1; i >= 0; i--)
            {
                try
                {
                    var window = asList[i];

                    if (window.Visible)
                    {
                        window.Close();
                        closedCount++;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Could not close window: {asList[i].Caption}, {ex.Message}");
                }
            }

            return closedCount;
        }
    }
}
