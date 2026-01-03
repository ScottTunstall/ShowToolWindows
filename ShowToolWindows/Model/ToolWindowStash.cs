using System;
using System.Linq;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Represents a saved snapshot of visible tool windows.
    /// </summary>
    public class ToolWindowStash
    {
        private string[] _windowCaptions;
        private string _description;

        /// <summary>
        /// Gets or sets the captions of the visible tool windows.
        /// </summary>
        public string[] WindowCaptions
        {
            get => _windowCaptions;
            set {
                _windowCaptions = value.Where(caption => !string.IsNullOrWhiteSpace(caption)).ToArray();
                _description = string.Join(", ", _windowCaptions);
            }
        }

        /// <summary>
        /// Gets or sets the DTE object kinds of the visible tool windows.
        /// </summary>
        public string[] WindowObjectKinds { get; set; }

        /// <summary>
        /// Gets or sets the timestamp when this stash was created.
        /// </summary>
        public DateTimeOffset CreatedAt { get; set; }

        public override string ToString()
        {
            return _description;
        }
    }
}
