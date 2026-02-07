using System.ComponentModel;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Wrapper class for displaying a ToolWindowStash with a dynamic index.
    /// </summary>
    public class StashListItem : INotifyPropertyChanged
    {
        private int _index;
        private ToolWindowStash _stash;

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

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public override string ToString()
        {
            return DisplayText;
        }
    }
}
