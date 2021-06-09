using LMFReports.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LMFReports.Services
{
    public enum ReportServiceType
    {
        Membership7447Service = 0,
        MembershipServiceNotValid = 1
    }
    internal interface IReportService
    {
        void AddWorkbook(string path);
        //void WriteToExcel(string path, List<IModelRequest> iModels);
        //List<IModelRequest> ReadFromExcel(string path);
    }
}
