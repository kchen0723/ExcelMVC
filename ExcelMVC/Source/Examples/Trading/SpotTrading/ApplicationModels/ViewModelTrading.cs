namespace SpotTrading.ApplicationModels
{
    using System.Linq;
    using BusinessModels;
    using CommandSinks;
    using ExcelMvc.Controls;
    using ExcelMvc.Views;
    using ViewModels;

    public class ViewModelTrading
    {
        public ViewModelTrading(View book)
        {
            // static ccy pair table (OneWayToSource)
            var tblCcyPair = (Table)book.Find("ExcelMvc.Table.CcyPairs");
            var pairs = new CcyPairs(tblCcyPair.MaxItemsToBind);
            tblCcyPair.Model = pairs;

            // static ccy list (OneWay)
            var tblCcys = book.Find("ExcelMvc.Table.Ccys");
            tblCcys.Model = pairs.Ccys;

            // exchange rates
            var tblRates = book.Find("ExcelMvc.Table.Rates");
            var rates = new ViewModelExchangeRates(new ExchangeRates(pairs));
            tblRates.Model = rates;

            // auto rate command
            var cmd = book.FindCommand("ExcelMvc.Command.AutoRate");
            cmd.Model = new CommandSinkAutoRate(rates);
            cmd.ClickedCaption = "Stop Simulation";

            // deal form
            var deal = new ViewModelDeal(rates);
            book.Find("ExcelMvc.Form.Deal").Model = deal;
            book.FindCommand("ExcelMvc.Command.InsideMode").Clicked += (x, y) =>
            {
                deal.IsInsideTrading = System.Convert.ToBoolean(((Command)x).Value);
            };

            // position table
            var tblPositions = (Table)book.Find("ExcelMvc.Table.Positions");
            var positions = new ViewModelPositions(tblPositions.MaxItemsToBind);
            tblPositions.Model = positions;
            book.FindCommand("ExcelMvc.Command.Reset").Clicked += (x, y) => positions.Reset();

            // manual deal command
            book.FindCommand("ExcelMvc.Command.ManualDeal").Model = new CommandSinkManualDeal(deal, positions, rates);

            var dealing = new ViewModelDealing(pairs.Ccys.ToList(), deal, positions, rates);
            cmd = book.FindCommand("ExcelMvc.Command.AutoDeal");
            cmd.Model = new CommandSinkAutoDeal(dealing);
            cmd.ClickedCaption = "Stop Auto-Deal";
        }
    }
}
