namespace LMFReports.Models.SheetRequests
{
    internal class ThreeYearMembershipReportRequest  //: IModelRequest
    {
        internal string DisplayName { get; set; }
        internal string PaymentDate { get; set; }
        internal string AmountReceived { get; set; }

        internal decimal AmountAsNumber
        {
            get 
            {
                decimal value;
                if (decimal.TryParse(AmountReceived.Replace("$","").Replace(",",""), out value))
                    return value;

                return 0;
            }

        }
        internal string FiscalYear { get; set; }

       //public int ColumnCount => 5; 
    }
}
