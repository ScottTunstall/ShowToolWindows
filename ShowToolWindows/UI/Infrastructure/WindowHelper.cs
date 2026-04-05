using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Helper class for managing Visual Studio tool windows.
    /// </summary>
    internal static class WindowHelper
    {
        /// <summary>
        /// Closes the specified collection of Visual Studio windows.
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

        /// <summary>
        /// Repositions a Visual Studio tool window if any part of it is off-screen.
        /// </summary>
        /// <param name="visualStudioWindow">The main Visual Studio window used to determine the vertical position.</param>
        /// <param name="windowToReposition">The tool window to reposition if off-screen.</param>
        /// <param name="screen">The virtual screen bounds.</param>
        public static void RepositionIfOffscreen(Window visualStudioWindow, Window windowToReposition, Rectangle screen)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            int winLeft = windowToReposition.Left;
            int winTop = windowToReposition.Top;
            int winWidth = windowToReposition.Width;
            int winHeight = windowToReposition.Height;
            int winRight = winLeft + winWidth;
            int winBottom = winTop + winHeight;

            var y = (visualStudioWindow.Top + visualStudioWindow.Height) > screen.Bottom
                ? screen.Top
                : visualStudioWindow.Top;

            if (winLeft < screen.Left)
            {
                FloatWindowAt(windowToReposition, screen.Left, y, winWidth, winHeight, screen);
                return;
            }

            if (winRight > screen.Right)
            {
                FloatWindowAt(windowToReposition, screen.Right - winWidth, y, winWidth, winHeight, screen);
                return;
            }

            if (winBottom > screen.Bottom)
            {
                FloatWindowAt(windowToReposition, winLeft, y, winWidth, winHeight, screen);
                return;
            }
        }

        private static void FloatWindowAt(Window window, int left, int top, int width, int height, Rectangle screen)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            window.IsFloating = true;

            int screenHeight = screen.Height;
            window.Width = Math.Min(width, 400);
            window.Height = Math.Min(height, screenHeight);
            window.Left = left;
            window.Top = top;
            window.Activate();
        }
    }
}