using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Collection of tool window stashes.
    /// </summary>
    public class StashCollection : ObservableCollection<ToolWindowStash>
    {
        /// <summary>
        /// Determines whether a stash with the same ordered object kinds already exists.
        /// </summary>
        /// <param name="objectKinds">The selected tool window object kinds.</param>
        /// <returns>True if a matching stash exists; otherwise, false.</returns>
        public bool HasMatchingConfiguration(IReadOnlyList<string> objectKinds)
        {
            return this.Any(stash => stash.MatchesSelection(objectKinds));
        }

        /// <summary>
        /// Creates a new stash from the provided captions and object kinds and inserts it at the top of the collection.
        /// </summary>
        /// <param name="captions">The tool window captions to include in the stash.</param>
        /// <param name="objectKinds">The tool window object kinds to include in the stash.</param>
        public void AddAndCreate(IEnumerable<string> captions, IEnumerable<string> objectKinds)
        {
            var stash = new ToolWindowStash
            {
                WindowCaptions = captions.ToArray(),
                WindowObjectKinds = objectKinds.ToArray(),
                CreatedAt = DateTimeOffset.UtcNow
            };

            Insert(0, stash);
        }

        public void DeleteTopOfStack()
        {
            RemoveAt(0);
        }
    }
}
