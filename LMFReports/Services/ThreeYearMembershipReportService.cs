using LMFReports.Models;
using LMFReports.Models.SheetRequests;
using LMFReports.Models.SheetResponses;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LMFReports.Services
{
    internal class ThreeYearMembershipReportService :IReportService
    {

        public void AddWorkbook(string path)
        {
            WriteToExcel(path, ReadFromExcel(path));

        }


        private void WriteToExcel(string path, List<ThreeYearMembershipReportRequest> iModels)
        {

            var requestModels = iModels.Cast<ThreeYearMembershipReportRequest>().ToList();
            var groupFys = GetFiscalYearGroups(requestModels);
            FileInfo filePath = new FileInfo(path);
            using (var excelPack = new ExcelPackage(filePath))
            {
                try 
                { 
                    var ws = excelPack.Workbook.Worksheets.Add("Three Year Membership Report");
                    WriteHeaderRow1(ws, groupFys);
                    WriteHeaderRow2(ws);
                    WriteDetailLines(ws, groupFys, requestModels);

                    excelPack.Save();
                }
                catch(Exception ex)
                {
                    Console.Write(ex.Message);
                   // Console.WriteLine("File has already been written to");
                    Console.ReadLine();
                }


            }
        }

        private void WriteDetailLines(ExcelWorksheet ws, ThreeMembershipReportResponse[] groupFys, List<ThreeYearMembershipReportRequest> requestModels)
        {
            var fy = requestModels.Select(y => y.FiscalYear).Distinct().ToArray();
            var displayNames = requestModels.Select(d=>d.DisplayName).Distinct();
            int row = 2;
            foreach (var displayName in displayNames)
            {
                var firstDate = requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[0]).FirstOrDefault()?.PaymentDate;
                var sum = $"${requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[0]).Sum(s => s.AmountAsNumber)}";

                row++;

                ws.Cells[$"A{row}:A{row}"].Value = displayName;
                ws.Cells[$"B{row}:B{row}"].Value = firstDate;
                ws.Cells[$"C{row}:C{row}"].Value = sum;

            }

            row = 2;
            foreach (var displayName in displayNames)
            {
                var firstDate = requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[1]).FirstOrDefault()?.PaymentDate;
                var sum = $"${requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[1]).Sum(s => s.AmountAsNumber)}";

                row++;

                ws.Cells[$"D{row}:D{row}"].Value = firstDate;
                ws.Cells[$"E{row}:E{row}"].Value = sum;

            }

            row = 2;
            foreach (var displayName in displayNames)
            {

                var firstDate = requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[2]).FirstOrDefault()?.PaymentDate;
                var sum = $"${requestModels.Where(rm => rm.DisplayName == displayName && rm.FiscalYear == fy[2]).Sum(s => s.AmountAsNumber)}";

                row++;

                ws.Cells[$"F{row}:F{row}"].Value = firstDate;
                ws.Cells[$"G{row}:G{row}"].Value = sum;

            }
        }

        private void WriteHeaderRow1(ExcelWorksheet ws, ThreeMembershipReportResponse[] groupFys)
        {
            ws.SelectedRange["A1:A1"].Value = "Display Name";
            ws.SelectedRange["A1:A1"].AutoFitColumns();
            ws.SelectedRange["B1:C1"].Value = $"FY {groupFys[0].Fy}";
            ws.SelectedRange["B1:C1"].Merge = true;
            ws.SelectedRange["D1:E1"].Value = $"FY {groupFys[1].Fy}";
            ws.SelectedRange["D1:E1"].Merge = true;
            ws.SelectedRange["F1:G1"].Value = $"FY {groupFys[2].Fy}";
            ws.SelectedRange["F1:G1"].Merge = true;

            ws.Row(1).Style.Font.Bold = true;
            ws.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

        }

        private void WriteHeaderRow2(ExcelWorksheet ws)
        {
            ws.Cells["B2:B2"].Value = "Payment Date";
            ws.SelectedRange["B2:B2"].AutoFitColumns();
            ws.Cells["C2:C2"].Value = "Amount";
            ws.SelectedRange["C2:C2"].AutoFitColumns();
            ws.Cells["D2:D2"].Value = "Payment Date";
            ws.SelectedRange["D2:D2"].AutoFitColumns();
            ws.Cells["E2:E2"].Value = "Amount";
            ws.SelectedRange["E2:E2"].AutoFitColumns();
            ws.Cells["F2:F2"].Value = "Payment Date";
            ws.SelectedRange["F2:F2"].AutoFitColumns();
            ws.Cells["G2:G2"].Value = "Amount";
            ws.SelectedRange["G2:G2"].AutoFitColumns();

            ws.Row(2).Style.Font.Bold = true;

        }

        private ThreeMembershipReportResponse[] GetFiscalYearGroups(List<ThreeYearMembershipReportRequest> requestModels)
        {
            return (from rm in requestModels
                    group rm by new
                    {
                        rm.FiscalYear,
                        //rm.DisplayName

                    }
                  into grm
                    select new ThreeMembershipReportResponse()
                    {
                        Fy = grm.Key.FiscalYear,
                        //DisplayNames = requestModels.Select(s=>s.DisplayName)
                        //.Distinct().ToList()
                    }).ToArray();

        }
        private List<ThreeYearMembershipReportRequest> ReadFromExcel(string path)
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excelPack = new ExcelPackage())
            {
                //Load excel stream
                using (var stream = File.OpenRead(path))
                {
                    excelPack.Load(stream);
                }

                //Lets Deal with first worksheet.(You may iterate here if dealing with multiple sheets)
                var ws = excelPack.Workbook.Worksheets[0];
                return GetRequests(ws); 
            }
        }

        private List<ThreeYearMembershipReportRequest> GetRequests(ExcelWorksheet ws)
        {
            //Get row details
            var startRow = 2;

            var membership7447Requests = new List<ThreeYearMembershipReportRequest>();

            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                
                membership7447Requests.Add(GetRequest(ws,rowNum));
            }
            
            return membership7447Requests;
        }

        private ThreeYearMembershipReportRequest GetRequest(ExcelWorksheet ws, int rowNum)
        {
            var membership7447Request = new ThreeYearMembershipReportRequest();
            var wsRow = ws.Cells[rowNum, 1, rowNum, 4];
            membership7447Request.DisplayName = wsRow[rowNum, 1].Text;
            membership7447Request.PaymentDate = wsRow[rowNum, 2].Text;
            membership7447Request.AmountReceived = wsRow[rowNum, 3].Text;
            membership7447Request.FiscalYear = wsRow[rowNum, 4].Text;

            return membership7447Request;
        }
    }
}
