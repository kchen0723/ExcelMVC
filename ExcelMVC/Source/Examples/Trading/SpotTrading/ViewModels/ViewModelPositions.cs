namespace SpotTrading.ViewModels
{
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using BusinessModels;

    public class ViewModelPositions :  List<ViewModelPosition>, INotifyCollectionChanged
    {
        public event NotifyCollectionChangedEventHandler CollectionChanged = delegate { };

        public ViewModelPositions(int count)
        {
            for (int idx = 0; idx < count; idx++)
                Add(new ViewModelPosition(new Position {Ccy1 = "", Ccy2 = ""}));
        }

        public void Net(Deal deal, ExchangeRates rates)
        {
            var item = this.FirstOrDefault(x => CcyPair.IsMatched(deal.BuyCcy, deal.SellCcy, x.Model.Ccy1, x.Model.Ccy2))
                       ?? this.FirstOrDefault(x => x.Model.Ccy1 == "");
            item.Net(deal, rates);
        }

        public void Reset()
        {
            foreach (var item in this)
                item.Model.Clear();
            CollectionChanged(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
}
