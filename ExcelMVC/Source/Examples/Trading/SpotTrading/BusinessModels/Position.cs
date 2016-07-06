
namespace SpotTrading.BusinessModels
{
    using System;

    public class Position
    {
        public string Ccy1 { get; set; }
        public string Ccy2 { get; set; }

        public double Amount1 { get; set; }
        public double Amount2 { get; set; }

        public double BaseAmount { get; set; }

        public void Net(Deal deal, ExchangeRates rates)
        {
            if (string.IsNullOrEmpty(Ccy1))
            {
                Ccy1 = deal.BuyCcy;
                Ccy2 = deal.SellCcy;

                Amount1 += deal.BuyAmount;
                Amount2 += -deal.SellAmount;
            }
            else
            {
                Amount1 += deal.BuyCcy == Ccy1 ? deal.BuyAmount : -deal.SellAmount;
                Amount2 += deal.BuyCcy == Ccy1 ? -deal.SellAmount : deal.BuyAmount;
            }

            Func<string, double, double> baseConversiion = (ccy, amount) =>
            {
                if (ccy == "USD")
                    return amount;

                var baseRate = rates.Find(ccy, "USD");
                if (baseRate.Pair.Ccy1 == ccy)
                    return amount * baseRate.Ask;

                return amount / baseRate.Bid;
            };

            BaseAmount = baseConversiion(Ccy1, Amount1);
            BaseAmount += baseConversiion(Ccy2, Amount2);
        }

        public void Clear()
        {
            Ccy1 = "";
            Ccy2 = "";
            Amount1 = 0;
            Amount2 = 0;
            BaseAmount = 0;
        }
    }
}
