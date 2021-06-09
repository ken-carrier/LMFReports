using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using LMFReports.Models.SheetRequests;

namespace LMFReports.Models.SheetResponses
{
    internal class Membership7447Response : IModelResponse
    {

        internal Fy Fy1 { get; init; }
        internal Fy Fy2 { get; init; }
        internal Fy Fy3 { get; init; }

        public string SheetName => "3-Year Membership Report";

      
    }

    internal class Fy
    {
       internal List<Payment> Payments { get; init; }
    }

    internal class Payment
    {
        internal List<string> DisplayName { get; init; }
        internal string PaymentDate { get; init; }
        internal string Amount { get; init; }

      
    }

    
        
}