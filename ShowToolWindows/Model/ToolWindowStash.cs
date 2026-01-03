using System;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Represents a saved snapshot of visible tool windows.
    /// </summary>
    public class ToolWindowStash
    {
        /// <summary>
        /// Gets or sets the captions of the visible tool windows.
        /// </summary>
        public string[] WindowCaptions { get; set; }

        /// <summary>
        /// Gets or sets the DTE object kinds of the visible tool windows.
        /// </summary>
        public string[] WindowObjectKinds { get; set; }

        /// <summary>
        /// Gets or sets the timestamp when this stash was created.
        /// </summary>
        public DateTime CreatedAt { get; set; }
    }
}
