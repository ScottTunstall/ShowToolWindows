using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Collection of tool window stashes.
    /// </summary>
    public class ToolWindowStashCollection : ObservableCollection<ToolWindowStash>
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
        /// Creates a new stash from the provided captions and object kinds, then adds it to the top of the collection.
        /// </summary>
        /// <param name="captions">The tool window captions to include in the stash.</param>
        /// <param name="objectKinds">The tool window object kinds to include in the stash.</param>
        /// <remarks>
        /// The <paramref name="captions"/> and <paramref name="objectKinds"/> sequences must have a 1:1 positional mapping.
        /// Each caption must correspond to the window kind at the same index (for example, caption at index 0 must be the caption of the tool window object kind at index 0).
        /// </remarks>
        public ToolWindowStash CreateAndPushToTop(IEnumerable<string> captions, IEnumerable<string> objectKinds)
        {
            var stash = new ToolWindowStash
            {
                WindowCaptions = captions.ToArray(),
                WindowObjectKinds = objectKinds.ToArray(),
                CreatedAt = DateTimeOffset.UtcNow
            };

            Insert(0, stash);
            return stash;
        }

        /// <summary>
        /// Removes the most recently added stash from the top of the collection.
        /// </summary>
        public void DeleteTopOfStack()
        {
            RemoveAt(0);
        }
    }
}
