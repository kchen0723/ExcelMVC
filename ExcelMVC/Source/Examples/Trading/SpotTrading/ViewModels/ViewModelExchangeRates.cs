namespace SpotTrading.ViewModels
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Specialized;
    using System.Linq;
    using System.Threading;
    using BusinessModels;

    public class ViewModelExchangeRates : List<ViewModelExchangeRate>, INotifyCollectionChanged
    {
        private ManualResetEvent AutoUpDateEvent { get; set; }
        public event NotifyCollectionChangedEventHandler CollectionChanged = delegate { };

        public ExchangeRates Model { get; private set; }

        public ViewModelExchangeRates(ExchangeRates rates)
        {
            Model = rates;
            Create();
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
            var random = new Random();
            while (!AutoUpDateEvent.WaitOne(2000))
            {
                var idx = (int)(random.NextDouble() * Count);
                if (idx >= Count) idx--;
                this[idx].Update();
            }
        }

        private void Create()
        {
            Clear();
            AddRange(Model.Select(x => new ViewModelExchangeRate { Model = x }));
            CollectionChanged(this, new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
        }
    }
}