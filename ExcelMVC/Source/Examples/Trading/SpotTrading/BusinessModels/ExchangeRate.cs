namespace SpotTrading.BusinessModels
{
    using System;

    public class ExchangeRate
    {
        public CcyPair Pair { get; set; }
        public double Bid { get; set; }
        public double Ask { get; set; }

        public ExchangeRate Flip()
        {
            var rate = new ExchangeRate { Pair = new CcyPair { Ccy1 = Pair.Ccy2, Ccy2 = Pair.Ccy1, Pip = Pair.Pip } };
            rate.Pair.Pip = Pair.Pip;
            rate.Bid = 1.0 / Bid;
            rate.Ask = 1.0 / Ask;
            return rate;
        }

        public static ExchangeRate Cross(ExchangeRate lhs, ExchangeRate rhs)
        {
            Func<CcyPair, string> nonBaseCcy = x => x.Ccy1 == "USD" ? x.Ccy2 : x.Ccy1;
            var rate = new ExchangeRate { Pair = new CcyPair { Ccy1 = nonBaseCcy(lhs.Pair), Ccy2 = nonBaseCcy(rhs.Pair) } };
            rate.Pair.Pip = Math.Max(lhs.Pair.Pip, rate.Pair.Pip);

            ExchangeRate ccy1Base;
            ExchangeRate ccy2Base;
            if (rate.Pair.Ccy1 == lhs.Pair.Ccy1 || rate.Pair.Ccy1 == lhs.Pair.Ccy2)
            {
                ccy1Base = lhs;
                ccy2Base = rhs;
            }
            else
            {
                ccy1Base = rhs;
                ccy2Base = lhs;
            }

            if (rate.Pair.Ccy1 == ccy1Base.Pair.Ccy2)
                ccy1Base = ccy1Base.Flip();

            if (rate.Pair.Ccy2 == ccy2Base.Pair.Ccy1)
                ccy2Base = rhs.Flip();

            rate.Bid = ccy1Base.Bid * ccy2Base.Ask;
            rate.Ask = ccy1Base.Ask * ccy2Base.Bid;

            return rate;
        }
    }
}
