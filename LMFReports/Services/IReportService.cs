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
        ThreeYearMembershipReportService = 0,
        MembershipServiceNotValid = 1
    }
    internal interface IReportService
    {
        void AddWorkbook(string path);
  
    }
}
