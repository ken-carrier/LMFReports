using System.Collections.Generic;

namespace LMFReports.Models.SheetResponses
{
    internal class ThreeMembershipReportResponse : IModelResponse
    {
        internal string Fy { get; set; }
        internal List<string> DisplayNames { get; set; }

     

    }
            
}