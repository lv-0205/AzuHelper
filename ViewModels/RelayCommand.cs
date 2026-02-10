using System;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace AzuHelper_v2.ViewModels;

public sealed class RelayCommand : ICommand
{
    private readonly Action? _execute;
    private readonly Func<Task>? _executeAsync;
    private readonly Func<bool>? _canExecute;

    public RelayCommand(Action execute, Func<bool>? canExecute = null)
    {
        _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        _canExecute = canExecute;
    }

    public RelayCommand(Func<Task> executeAsync, Func<bool>? canExecute = null)
    {
        _executeAsync = executeAsync ?? throw new ArgumentNullException(nameof(executeAsync));
        _canExecute = canExecute;
    }

    public bool CanExecute(object? parameter) => _canExecute?.Invoke() ?? true;

    public async void Execute(object? parameter)
    {
        if (_executeAsync is not null)
        {
            await _executeAsync().ConfigureAwait(false);
            return;
        }

        _execute?.Invoke();
    }

    public event EventHandler? CanExecuteChanged;

    public void RaiseCanExecuteChanged()
    {
        _ = Application.Current.Dispatcher.InvokeAsync(() => CanExecuteChanged?.Invoke(this, EventArgs.Empty));
    }
}
