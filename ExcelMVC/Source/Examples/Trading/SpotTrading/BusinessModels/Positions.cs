
namespace SpotTrading.BusinessModels
{
    using System.Collections.Generic;
    using System.Linq;

    public class Positions : List<Position>
    {
        public void Net(Deal deal, ExchangeRates rates)
        {
            var item = this.FirstOrDefault(x => CcyPair.IsMatched(x.Ccy1, x.Ccy2, deal.BuyCcy, deal.SellCcy));
            if (item == null)
            {
                item = new Position();
                Add(item);
            }
            item.Net(deal, rates);
        }
    }
}
