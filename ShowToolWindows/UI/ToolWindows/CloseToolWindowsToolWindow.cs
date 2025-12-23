using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace ShowToolWindows.UI.ToolWindows
{
    /// <summary>
    /// Tool window that hosts the UI for showing and hiding tool windows.
    /// </summary>
    [Guid(ToolWindowGuidString)]
    public sealed class CloseToolWindowsToolWindow : ToolWindowPane
    {
        /// <summary>
        /// Tool window GUID.
        /// </summary>
        public const string ToolWindowGuidString = "9c6ea6bd-2b57-4ff3-b429-3caa28b3e31c";

        internal const string ToolWindowTitle = "Show/Hide specific tool windows";

        /// <summary>
        /// Initializes a new instance of the <see cref="CloseToolWindowsToolWindow"/> class.
        /// </summary>
        public CloseToolWindowsToolWindow()
            : base(null)
        {
            Caption = ToolWindowTitle;
            Content = new CloseToolWindowsControl();
        }

        /// <summary>
        /// Initializes the tool window content with package services.
        /// </summary>
        /// <param name="package">The owning package.</param>
        public async Task InitializeAsync(AsyncPackage package)
        {
            if (package == null)
            {
                throw new ArgumentNullException(nameof(package));
            }

            CloseToolWindowsControl control = (CloseToolWindowsControl)Content;
            await control.InitializeAsync(package);
        }
    }
}
