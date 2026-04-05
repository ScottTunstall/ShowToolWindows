using ShowToolWindows.Model;
using System.ComponentModel;

namespace ShowToolWindows.UI.ViewModels
{
    /// <summary>
    /// Wrapper class for displaying a ToolWindowStash with a dynamic index.
    /// </summary>
    public class StashListItem : INotifyPropertyChanged
    {
        private int _index;
        private ToolWindowStash _stash;

        /// <summary>
        /// Initializes a new instance of the <see cref="StashListItem"/> class.
        /// </summary>
        /// <param name="stash">The stash to display.</param>
        /// <param name="index">The display index for the stash.</param>
        public StashListItem(ToolWindowStash stash, int index)
        {
            _stash = stash;
            _index = index;
        }

        /// <summary>
        /// Gets or sets the underlying stash object.
        /// </summary>
        public ToolWindowStash Stash
        {
            get => _stash;
            set
            {
                if (_stash != value)
                {
                    _stash = value;
                    OnPropertyChanged(nameof(Stash));
                    OnPropertyChanged(nameof(DisplayText));
                }
            }
        }

        /// <summary>
        /// Gets the formatted display text with index and description.
        /// </summary>
        public string DisplayText => $"{{{_index}}} - {_stash}";

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// Returns the display text for the stash list item.
        /// </summary>
        /// <returns>The formatted display text.</returns>
        public override string ToString()
        {
            return DisplayText;
        }
    }
}