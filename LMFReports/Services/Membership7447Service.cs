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
    internal class Membership7447Service :IReportService
    {

        public void AddWorkbook(string path)
        {
            WriteToExcel(path, ReadFromExcel(path));

        }


        private void WriteToExcel(string path, List<Membership7447Request> iModels)
        {
            //Let use below test data for writing it to excel
            // let's convert our object data to Datatable for a simplified logic.
            // Datatable is the easiest way to deal with complex datatypes for easy reading and formatting. 

            var requestModels = iModels.Cast<Membership7447Request>().ToList();
            var groupFys= (from rm in requestModels
                              group rm by new
                              {
                                  rm.FiscalYear,
                                  // rm.DisplayName

                              }
                              into grm
                              select new Fy()
                              {
                                  Year = grm.Key.FiscalYear,
                                  Payments = grm.ToList()

                                                                // Fy1 = new Fy() { new List<Payment>() }; // { Amount = grm.Sum(s=>s.AmountAsLong).ToString()}
                              }).ToArray();

            //var groupFys = (from gdn in groupDispayNames
            //                       group gdn by new
            //                       {
            //                          gdn.Year
            //                           // rm.DisplayName

            //                       }
            //                  into grds
            //                       select new Membership7447Response()
            //                       {
            //                           //Fy1 = grds.Where(w=>w.Year == grds.Key.Year && grds.Where(a=>a.DisplayName =v))
            //                           //Sum =  

            //                           // Fy1 = new Fy() { new List<Payment>() }; // { Amount = grm.Sum(s=>s.AmountAsLong).ToString()}
            //                       }).ToArray();


            //foreach (var fy in groupDispayNames)
            //{

            //}

            //DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(requestModels), (typeof(DataTable)));
            FileInfo filePath = new FileInfo(path);
            using (var excelPack = new ExcelPackage(filePath))
            {
                var ws = excelPack.Workbook.Worksheets.Add("3-Year Membership Report");
                var wsRow = ws.Cells[1, 1, 1, 7];
                ws.SelectedRange["A1:A1"].Value = "Display Name";
                ws.SelectedRange["B1:C1"].Value = $"FY {groupFys[0].Year}";
                ws.SelectedRange["B1:C1"].Merge = true;
                ws.SelectedRange["D1:E1"].Value = $"FY {groupFys[1].Year}";
                ws.SelectedRange["F1:G1"].Value = $"FY {groupFys[2].Year}";
                ws.SelectedRange["D1:E1"].Merge = true;
                ws.SelectedRange["F1:G1"].Merge = true;

                ws.Row(1).Style.Font.Bold = true;
                ws.Row(1).Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                ws.Cells["B2:B2"].Value = "Payment Date";
                ws.Cells["C2:C2"].Value = "Amount";
                ws.Cells["D2:D2"].Value = "Payment Date";
                ws.Cells["E2:E2"].Value = "Amount";
                ws.Cells["F2:F2"].Value = "Payment Date";
                ws.Cells["G2:G2"].Value = "Amount";

                ws.Row(2).Style.Font.Bold = true;

                excelPack.Save();
            }
        }

        private List<Membership7447Request> ReadFromExcel(string path)
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

        private List<Membership7447Request> GetRequests(ExcelWorksheet ws)
        {
            //Get row details
            var startRow = 2;

            var membership7447Requests = new List<Membership7447Request>();

            for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
            {
                
                membership7447Requests.Add(GetRequest(ws,rowNum));
            }
            
            return membership7447Requests;
        }

        private Membership7447Request GetRequest(ExcelWorksheet ws, int rowNum)
        {
            var membership7447Request = new Membership7447Request();
            var wsRow = ws.Cells[rowNum, 1, rowNum, 4];
            membership7447Request.DisplayName = wsRow[rowNum, 1].Text;
            membership7447Request.PaymentDate = wsRow[rowNum, 2].Text;
            membership7447Request.AmountReceived = wsRow[rowNum, 3].Text;
            membership7447Request.FiscalYear = wsRow[rowNum, 4].Text;

            return membership7447Request;
        }
    }
}
