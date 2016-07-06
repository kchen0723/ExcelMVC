using System;
using System.ComponentModel;

namespace SpotTrading.ViewModels
{
    using BusinessModels;

    public class ViewModelExchangeRate : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public ExchangeRate Model { get; set; }

        public string Code
        {
            get { return Model.Pair.Code; }
        }

        public double Bid
        {
            get { return Model.Bid; }
        }

        public double Ask
        {
            get { return Model.Ask; }
        }

        public void RaiseChanged()
        {
            PropertyChanged(this, new PropertyChangedEventArgs("Bid"));
            PropertyChanged(this, new PropertyChangedEventArgs("Ask"));
        }

        public void Update()
        {
            var random = new Random();
            var change = (0.5 - random.NextDouble()) * Model.Pair.Pip * 10;
            var bid = Bid - change;
            var ask = Ask + change;
            Model.Bid = Math.Min(bid, ask);
            Model.Ask = Math.Max(bid, ask);
            RaiseChanged();
        }
    }
}
