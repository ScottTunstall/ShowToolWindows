using System;
using System.Collections.Generic;
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

        /// <summary>
        /// Determines whether this stash matches the provided ordered object kind selection exactly.
        /// </summary>
        /// <param name="objectKinds">The selected tool window object kinds.</param>
        /// <returns>True if object kinds match exactly in length and order; otherwise, false.</returns>
        public bool MatchesSelection(IReadOnlyList<string> objectKinds)
        {
            if (WindowObjectKinds == null || objectKinds == null)
            {
                return false;
            }

            if (WindowObjectKinds.Length != objectKinds.Count)
            {
                return false;
            }

            for (int i = 0; i < objectKinds.Count; i++)
            {
                if (!string.Equals(WindowObjectKinds[i], objectKinds[i], StringComparison.Ordinal))
                {
                    return false;
                }
            }

            return true;
        }

        public override string ToString()
        {
            return _description;
        }
    }
}
