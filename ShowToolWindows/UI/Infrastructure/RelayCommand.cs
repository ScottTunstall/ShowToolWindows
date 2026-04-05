using System;
using System.Windows.Input;

namespace ShowToolWindows.UI.Infrastructure
{
    /// <summary>
    /// Simple relay command implementation for binding keyboard shortcuts.
    /// </summary>
    internal class RelayCommand : ICommand
    {
        private readonly Action<object> _execute;

        /// <summary>
        /// Initializes a new instance of the <see cref="RelayCommand"/> class.
        /// </summary>
        /// <param name="execute">The action to invoke when the command is executed.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="execute"/> is null.</exception>
        public RelayCommand(Action<object> execute)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        }

        /// <summary>
        /// Occurs when changes occur that affect whether the command should execute.
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        /// <summary>
        /// Determines whether the command can execute.
        /// </summary>
        /// <param name="parameter">The command parameter.</param>
        /// <returns><c>true</c> in all cases.</returns>
        public bool CanExecute(object parameter) => true;

        /// <summary>
        /// Executes the command.
        /// </summary>
        /// <param name="parameter">The command parameter.</param>
        public void Execute(object parameter) => _execute(parameter);
    }
}