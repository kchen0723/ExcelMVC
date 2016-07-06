namespace SpotTrading.CommandSinks
{
    using System;
    using System.Windows.Input;
    using ExcelMvc.Controls;
    using ViewModels;

    public class CommandSinkAutoRate :  ICommand
    {
        private ViewModelExchangeRates Model { get; set; }
        private bool IsRunning { get; set; }

        public CommandSinkAutoRate(ViewModelExchangeRates rates)
        {
            Model = rates;
        }

        public bool CanExecute(object parameter)
        {
            // toggling betweeen start and stop
            return true;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
            ExecuteAutoRate();
        }

        private void ExecuteAutoRate()
        {
            IsRunning = !IsRunning;
            if (IsRunning)
                Model.StartSimulate();
            else
                Model.StopSimulate();
        }
    }
}