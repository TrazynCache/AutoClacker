using System;
using System.Threading.Tasks;
using System.Windows.Input;

namespace AutoClacker.Utilities
{
    public class RelayCommand : ICommand
    {
        private readonly Func<object, Task> executeAsync;
        private readonly Action<object> execute;
        private readonly Func<object, bool> canExecute;

        public RelayCommand(Func<object, Task> executeAsync, Func<object, bool> canExecute = null)
        {
            this.executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
            this.canExecute = canExecute;
        }

        public RelayCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
            this.canExecute = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter) => canExecute == null || canExecute(parameter);

        public async void Execute(object parameter)
        {
            if (CanExecute(parameter))
            {
                if (executeAsync != null)
                {
                    await executeAsync(parameter);
                }
                else
                {
                    execute(parameter); 
                }
            }
        }
    }
}
