namespace SpotTrading.BusinessModels
{
    public class CcyPair
    {
        public string Ccy1 { get; set; }
        public string Ccy2 { get; set; }
        public double Pip { get; set; }
        public double Spot { get; set; }

        public string Code
        {
            get { return string.Format("{0}/{1}", Ccy1, Ccy2); }
        }

        public bool IsValid
        {
            get { return !string.IsNullOrEmpty(Ccy1) && !string.IsNullOrEmpty(Ccy2); }
        }

        public bool IsMatched(string ccy1, string ccy2)
        {
            return IsMatched (ccy1, ccy2, Ccy1, Ccy2);
        }

        public static bool IsMatched(string ccy1, string ccy2, string xccy1, string xccy2)
        {
            return (ccy1 == xccy1 && ccy2 == xccy2) || (ccy1 == xccy2 && ccy2 == xccy1);
        }
    }
}

