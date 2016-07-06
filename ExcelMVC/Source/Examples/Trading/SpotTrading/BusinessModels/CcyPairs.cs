namespace SpotTrading.BusinessModels
{
    using System.Collections.Generic;
    using System.Linq;

    public class CcyPairs :List<CcyPair>
    {
        public CcyPairs(int count)
        {
            Create(count);
        }

        public void Create(int count)
        {
            Clear();
            while (--count >=0)
                Add(new CcyPair());
        }

        public IEnumerable<string> Ccys
        {
            get
            {
                var ccys = this.Select(y => y.Ccy1).ToList();
                ccys.AddRange(this.Select(y => y.Ccy2));
                ccys = ccys.Where(x => !string.IsNullOrEmpty(x)).Distinct().ToList();
                ccys.Sort();
                return ccys;
            }
        }
    }
}
