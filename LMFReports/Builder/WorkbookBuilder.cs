using LMFReports.Models;
using LMFReports.Services;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LMFReports.Builder
{
    internal interface IWorkbookBuilder
    {
        void BuildWorkbooks();

    }

    internal class WorkbookBuilder : IWorkbookBuilder
    {

        private Func<ReportServiceType, IReportService> _modelReportServiceDelegate { get; }
        public WorkbookBuilder(Func<ReportServiceType, IReportService> modelReportServiceDelegate) 
        {
            _modelReportServiceDelegate = modelReportServiceDelegate;
        }

        public void BuildWorkbooks()
        {

            string path = "Copy of 2021-0522 Membership Report(7447).xlsx";

            IReportService reportService = _modelReportServiceDelegate(GetReportServiceType(path));
            reportService.AddWorkbook(path);
            
        }

        ReportServiceType GetReportServiceType(string path)
        {
            return (path.Contains("Membership Report(7447)")) switch
            {
                (true) => ReportServiceType.Membership7447Service,
                _ => ReportServiceType.MembershipServiceNotValid
            };
        }

    }
}
