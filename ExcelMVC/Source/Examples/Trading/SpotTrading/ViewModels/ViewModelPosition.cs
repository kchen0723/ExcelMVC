namespace SpotTrading.ViewModels
{
    using System.ComponentModel;
    using BusinessModels;

    public class ViewModelPosition : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { }; 

        public Position Model { get; private set; }

        public ViewModelPosition(Position model)
        {
            Model = model;
        }

        public void Net(Deal deal, ExchangeRates rates)
        {
            bool everything = string.IsNullOrEmpty(Model.Ccy1);
            Model.Net(deal, rates);
            RaiseChanged(everything);
        }

        public void RaiseChanged(bool everything)
        {
            if (everything)
            {
                PropertyChanged(this, new PropertyChangedEventArgs("Model.Ccy1"));
                PropertyChanged(this, new PropertyChangedEventArgs("Model.Ccy2"));
            }
            PropertyChanged(this, new PropertyChangedEventArgs("Model.Amount1"));
            PropertyChanged(this, new PropertyChangedEventArgs("Model.Amount2"));
            PropertyChanged(this, new PropertyChangedEventArgs("Model.BaseAmount"));
        }
    }
}
