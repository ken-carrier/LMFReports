namespace LMFReports.Models.SheetRequests
{
    internal class Membership7447Request  //: IModelRequest
    {
        internal string DisplayName { get; set; }
        internal string PaymentDate { get; set; }
        internal string AmountReceived { get; set; }

        internal long AmountAsLong
        {
            get 
            {
                long value;
                if (long.TryParse(AmountReceived, out value))
                    return value;

                return 0;
            }

        }
        internal string FiscalYear { get; set; }

       //public int ColumnCount => 5; 
    }
}
