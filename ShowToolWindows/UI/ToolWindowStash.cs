using System;

namespace ShowToolWindows.UI
{
    /// <summary>
    /// Represents a saved snapshot of visible tool windows.
    /// </summary>
    public class ToolWindowStash
    {
        public string[] WindowCaptions { get; set; }

        /// <summary>
        /// Gets or sets the pipe-separated list of window ObjectKind GUIDs.
        /// </summary>
        public string[] WindowObjectKinds { get; set; }

        /// <summary>
        /// Gets or sets the timestamp when this stash was created.
        /// </summary>
        public DateTime CreatedAt { get; set; }
    }
}
