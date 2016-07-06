namespace SpotTrading.ViewModels
{
    using System.ComponentModel;
    using BusinessModels;

    public class ViewModelDeal : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = delegate { };

        public Deal Model { get; private set; }
        public ViewModelExchangeRates Rates { get; private set; }
        public bool IsInsideTrading { get; set; }

        public ViewModelDeal(ViewModelExchangeRates rates)
        {
            Model = new Deal();
            Rates = rates;
            foreach (var rate in Rates)
                rate.PropertyChanged += rate_PropertyChanged;
        }

        void rate_PropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            var rate = (ViewModelExchangeRate) sender;
            if (ReferenceEquals(rate.Model, Model.Rate))
                SetRate(rate.Model);
        }

        public string BuyCcy
        {
            get 
            {
                return Model.BuyCcy; 
            }
            set
            {
                Model.BuyCcy = value.ToUpper();
                SetRate();
            }
        }

        public string SellCcy
        {
            get
            {
                return Model.SellCcy;
            }
            set
            {
                Model.SellCcy = value.ToUpper();
                SetRate();
            }
        }


        public double BuyAmount
        {
            get
            {
                return Model.BuyAmount;
            }
            set
            {
                Model.BuyAmount = value;
                Model.IsCcy1Fixed = true;
                DeriveXAmount();
            }
        }

        public double SellAmount
        {
            get
            {
                return Model.SellAmount;
            }
            set
            {
                Model.SellAmount = value;
                Model.IsCcy1Fixed = false;
                DeriveXAmount();
            }
        }

        public void SetRate()
        {
            var fx = Rates.Model.Find(BuyCcy, SellCcy);
            if (fx == null)
                return;
            SetRate(fx);
        }

        private void SetRate(ExchangeRate fx)
        {
            Model.Rate = fx;
            DeriveXAmount();
        }

        public void DeriveXAmount()
        {
            if (Model.TryDeriveXAmount(IsInsideTrading))
                RaiseChanged(Model.IsCcy1Fixed ? "SellAmount" : "BuyAmount");
        }

        public void RaiseChanged(string name)
        {
            PropertyChanged(this, new PropertyChangedEventArgs((name)));
        }

        public void RaiseChanged()
        {
            RaiseChanged("BuyCcy");
            RaiseChanged("BuyAmount");
            RaiseChanged("SellCcy");
            RaiseChanged("SellAmount");
        }
    }
}
