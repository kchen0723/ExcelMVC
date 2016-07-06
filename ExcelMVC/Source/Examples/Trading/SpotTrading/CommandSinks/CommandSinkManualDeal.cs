namespace SpotTrading.CommandSinks
{
    using System;
    using System.Windows.Input;
    using ViewModels;

    public class CommandSinkManualDeal:  ICommand
    {
        private ViewModelDeal Deal { get; set; }
        private ViewModelPositions Positions { get; set; }
        private ViewModelExchangeRates Rates { get; set; }

        public CommandSinkManualDeal(ViewModelDeal deal, ViewModelPositions positions, ViewModelExchangeRates rates)
        {
            Deal = deal;
            Deal.PropertyChanged += Deal_PropertyChanged;
            Positions = positions;
            Rates = rates;
            CanExecuteChanged(this, null);
        }

        void Deal_PropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            CanExecuteChanged(this, null);
        }

        public bool CanExecute(object parameter)
        {
            return Deal.Model.BuyCcy != null
                && Deal.Model.SellCcy != null
                && Deal.Model.BuyCcy != Deal.Model.SellCcy;
        }

        public event EventHandler CanExecuteChanged = delegate { };

        public void Execute(object parameter)
        {
             Positions.Net(Deal.Model, Rates.Model);
        }
    }
}