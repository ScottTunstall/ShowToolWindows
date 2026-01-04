using EnvDTE;
using Microsoft.VisualStudio.Shell;
using System;
using System.ComponentModel;
using System.Diagnostics;

namespace ShowToolWindows.Model
{
    /// <summary>
    /// Represents a Visual Studio tool window entry that can be shown or hidden.
    /// </summary>
    public sealed class ToolWindowEntry : INotifyPropertyChanged
    {
        private readonly Window _window;
        private bool _isVisible;

        /// <summary>
        /// Initializes a new instance of the <see cref="ToolWindowEntry"/> class.
        /// </summary>
        /// <param name="window">The Visual Studio tool window.</param>
        public ToolWindowEntry(Window window)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            _window = window ?? throw new ArgumentNullException(nameof(window));
            Caption = window.Caption;
            ObjectKind = window.ObjectKind;
            _isVisible = window.Visible;
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
        /// Gets or sets a value indicating whether the tool window is visible.
        /// </summary>
        public bool IsVisible
        {
            get
            {
                return _isVisible;
            }
            set
            {
                if (_isVisible == value)
                {
                    return;
                }

                _isVisible = value;
                OnPropertyChanged(nameof(IsVisible));
            }
        }

        /// <summary>
        /// Occurs when a property value changes.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        /// <summary>
        /// Applies visibility changes to the underlying Visual Studio window.
        /// </summary>
        /// <param name="isVisible">Whether the window should be visible.</param>
        public void SetVisibility(bool isVisible)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (isVisible)
                {
                    _window.Visible = true;
                    _window.Activate();
                }
                else
                {
                    _window.Visible = false;
                }

                IsVisible = _window.Visible;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Unable to change visibility for window '{Caption}': {ex.Message}");
            }
        }

        /// <summary>
        /// Syncs the state with the underlying Visual Studio window.
        /// </summary>
        public void Synchronize()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            IsVisible = _window.Visible;
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
