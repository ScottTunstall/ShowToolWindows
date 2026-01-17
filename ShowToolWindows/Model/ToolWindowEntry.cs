using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Diagnostics;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Represents a Visual Studio tool window entry that can be selected for stashing.
    /// </summary>
    public sealed class ToolWindowEntry : INotifyPropertyChanged
    {
        private bool _isSelected;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowEntry"/> class.
        /// </summary>
        /// <param name="window">The Visual Studio tool window.</param>
        public ToolWindowEntry(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Caption = window.Caption;
            ObjectKind = window.ObjectKind;
            _isSelected = false;
        }

        /// <summary>
        /// Gets the display caption for the tool window.
        /// </summary>
        public string Caption
        {
            get;
        }

        /// <summary>
        /// Gets the object kind GUID for the tool window.
        /// </summary>
        public string ObjectKind
        {
            get;
        }

        /// <summary>
        /// Gets or sets a value indicating whether the tool window is selected for stashing.
        /// </summary>
        public bool IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                if (_isSelected == value)
                {
                    return;
                }

                _isSelected = value;
                OnPropertyChanged(nameof(IsSelected));
            }
        }

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Determines whether the specified object is equal to the current object.
        /// </summary>
        /// <param name="obj">The object to compare with the current object.</param>
        /// <returns>true if the specified object is equal to the current object; otherwise, false.</returns>
        public override bool Equals(object obj)
        {
            if (obj is ToolWindowEntry other)
            {
                return string.Equals(ObjectKind, other.ObjectKind, StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        /// <summary>
        /// Returns a hash code for the current object.
        /// </summary>
        /// <returns>A hash code for the current object.</returns>
        public override int GetHashCode()
        {
            return StringComparer.OrdinalIgnoreCase.GetHashCode(ObjectKind ?? string.Empty);
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handlers = PropertyChanged;
            if (handlers != null)
            {
                handlers(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
