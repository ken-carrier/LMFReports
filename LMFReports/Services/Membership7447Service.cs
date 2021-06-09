using LMFReports.Models;
using LMFReports.Models.SheetRequests;
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

            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(requestModels), (typeof(DataTable)));
            FileInfo filePath = new FileInfo(path);
            using (var excelPack = new ExcelPackage(filePath))
            {
                var ws = excelPack.Workbook.Worksheets.Add("WriteTest");
                ws.Cells.LoadFromDataTable(table, true, OfficeOpenXml.Table.TableStyles.Light8);
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
