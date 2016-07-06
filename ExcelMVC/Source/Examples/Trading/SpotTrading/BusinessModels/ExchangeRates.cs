namespace SpotTrading.BusinessModels
{
    using System.Collections.Generic;
    using System.Linq;

    public class ExchangeRates : List<ExchangeRate>
    {
        public ExchangeRates(IEnumerable<CcyPair> pairs)
        {
            Create(pairs);
        }

        public void Create(IEnumerable<CcyPair> pairs)
        {
            Clear();
            AddRange(pairs.Where(x => x.IsValid).Select(y => new ExchangeRate { Pair = y, Bid = y.Spot - y.Pip, Ask = y.Spot + y.Pip }));
        }

        public ExchangeRate Find(string ccy1, string ccy2)
        {
            if (ccy1 == null || ccy2 == null || ccy1 == ccy2)
                return null;

            var rate = this.FirstOrDefault(x => x.Pair.IsMatched(ccy1, ccy2));

            if (rate != null)
                return rate;

            var lhs = Find(ccy1, "USD");
            var rhs = Find(ccy2, "USD");

            return ExchangeRate.Cross(lhs, rhs);
        }
    }
}
