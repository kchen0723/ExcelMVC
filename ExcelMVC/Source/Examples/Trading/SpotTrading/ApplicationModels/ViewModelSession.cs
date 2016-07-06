namespace SpotTrading.ApplicationModels
{
    using ExcelMvc.Extensions;
    using ExcelMvc.Runtime;
    using ExcelMvc.Views;

    public class ViewModelSession : ISession
    {
        private const string BookId = "SpotTrading";
        public ViewModelSession()
        {
            // hook notificaton events
            App.Instance.Opening += Instance_Opening;
            App.Instance.Opened += Instance_Opened;
            App.Instance.Closing += Instance_Closing;
            App.Instance.Closed += Instance_Closed;
        }

        void Instance_Opening(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // accept view
                args.Accept();
            }
        }

        void Instance_Opened(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // assign model
                args.Accept();
                args.View.Model = new ViewModelTrading(args.View);
            }
        }

        void Instance_Closing(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // allow closing
                args.Accept();
            }
        }

        void Instance_Closed(object sender, ViewEventArgs args)
        {
            if (IsMybook(args))
            {
                // detach model
                args.View.Model = null;
            }
        }

        private bool IsMybook(ViewEventArgs args)
        {
            return args.View.Id.CompareOrdinalIgnoreCase(BookId) == 0;
        }

        public void Dispose()
        {
        }
    }
}
