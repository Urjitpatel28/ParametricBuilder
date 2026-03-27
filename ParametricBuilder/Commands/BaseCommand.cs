using NLog;
using System;
using System.Windows.Input;

namespace ParametricBuilder.Commands
{
    public class BaseCommand : ICommand
    {
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public BaseCommand()
        {
            CommandManager.RequerySuggested += OnRequerySuggested;
        }

        ~BaseCommand()
        {
            CommandManager.RequerySuggested -= OnRequerySuggested;
        }

        private void OnRequerySuggested(object sender, EventArgs e)
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }

        public event EventHandler CanExecuteChanged;

        public virtual bool CanExecute(object parameter)
        {
            // Default implementation always allows execution
            return true;
        }

        public virtual void Execute(object parameter)
        {
            // Default implementation does nothing
            // Override in derived class to provide specific functionality
        }
    }
}