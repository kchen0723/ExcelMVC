namespace SpotTrading.BusinessModels
{
    using System;

    public class Deal
    {
        public string BuyCcy { get; set; }
        public string SellCcy { get; set; }
        public double BuyAmount { get; set; }
        public double SellAmount { get; set; }
        public bool IsCcy1Fixed { get; set; }

        public ExchangeRate Rate { get; set; }

        public bool TryDeriveXAmount(bool inside)
        {
            if (Rate == null)
                return false;

            double fx;
            if (Rate.Pair.Ccy1 == BuyCcy)
                fx = inside ? Rate.Bid : Rate.Ask;
            else
                fx = 1 / (inside ? Rate.Ask : Rate.Bid);

            if (IsCcy1Fixed)
            {
                SellAmount = BuyAmount * fx;
                return true;
            }
            else if (Math.Abs(fx) > 0.000001)
            {
                BuyAmount = SellAmount / fx;
                return true;
            }
            return false;
        }
    }
}
