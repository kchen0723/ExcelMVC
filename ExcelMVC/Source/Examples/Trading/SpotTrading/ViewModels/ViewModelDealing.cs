namespace SpotTrading.ViewModels
{
    using System;
    using System.Collections.Generic;
    using System.Threading;

    public class ViewModelDealing
    {
        private ManualResetEvent AutoUpDateEvent { get; set; }

        private List<string> Ccys { get; set; }
        private ViewModelDeal Deal { get; set; }
        private ViewModelPositions Positions { get; set; }
        private ViewModelExchangeRates Rates { get; set; }

        public ViewModelDealing(List<string> ccys,ViewModelDeal deal, ViewModelPositions positions,  ViewModelExchangeRates rates)
        {
            Ccys = ccys;
            Deal = deal;
            Positions = positions;
            Rates = rates;
        }

        public void StartSimulate()
        {
            AutoUpDateEvent = new ManualResetEvent(false);
            var thread = new Thread(Update) { Name = "ExcelMvcAsynUpdateThread", IsBackground = true };
            thread.Start();
        }

        public void StopSimulate()
        {
            if (AutoUpDateEvent != null)
                AutoUpDateEvent.Set();
        }

        private void Update(object state)
        {
            while (!AutoUpDateEvent.WaitOne(2000))
            {
                MaketDeal();
                Positions.Net(Deal.Model, Rates.Model);
            }
        }

        private void MaketDeal()
        {
            var random = new Random();

            var idx = (int)(random.NextDouble() * Ccys.Count);
            if (idx >= Ccys.Count) idx = 0;

            var jdx = (int)(random.NextDouble() * Ccys.Count);
            if (jdx >= Ccys.Count) jdx = 0;

            if (jdx == idx)
                jdx = idx + 1;
            if (jdx >= Ccys.Count) jdx = 0;

            Deal.BuyCcy = Ccys[idx];
            Deal.SellCcy = Ccys[jdx];
            Deal.BuyAmount = (random.NextDouble() + 0.1)*1000000;
            Deal.RaiseChanged();
        }
    }
}