using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Windows = EnvDTE.Windows;

namespace ShowToolWindows
{
    internal static class WindowHelper
    {
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
                    window.Close();
                    closedCount++;
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Could not close window: {asList[i].Caption}, {ex.Message}");
                }
            }

            return closedCount;
        }

        public static IReadOnlyCollection<Window> GetAllToolWindowsExceptMainWindow(IEnumerable windows)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var toolWindows = new List<Window>();

            foreach (Window window in windows)
            {
                if (window.Kind != WindowKindConsts.ToolWindowKind)
                {
                    continue;
                }

                if (window.ObjectKind == Constants.vsWindowKindMainWindow)
                {
                    continue;
                }

                toolWindows.Add(window);
            }

            return toolWindows;
        }


    }
}
