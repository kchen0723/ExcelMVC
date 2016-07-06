namespace SpotTrading.CommandSinks
{
    using System;
    using System.Windows.Input;
    using ViewModels;

    public class CommandSinkAutoDeal :  ICommand
    {
        private ViewModelDealing Model { get; set; }
        private bool IsRunning { get; set; }

        public CommandSinkAutoDeal(ViewModelDealing deals)
        {
            Model = deals;
            CanExecuteChanged(this, new EventArgs());
        }

        public bool CanExecute(object parameter)
        {
            // toggling betweeen start and stop
            return true;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            ExecuteAutoDeal();
        }

        private void ExecuteAutoDeal()
        {
            IsRunning = !IsRunning;
            if (IsRunning)
                Model.StartSimulate();
            else
                Model.StopSimulate();
        }
    }
}