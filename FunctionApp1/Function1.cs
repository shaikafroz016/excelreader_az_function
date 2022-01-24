using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net;
using ExcelDataReader;
using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace FunctionApp1
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static string Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            log.LogInformation($"C# HTTP trigger function processed a request.");

            WebClient client = new WebClient();

            byte[] buffer = client.DownloadData("https://teststaccshaik.blob.core.windows.net/excel/Book1.xlsx");

            MemoryStream stream = new MemoryStream();
            stream.Write(buffer, 0, buffer.Length);
            stream.Position = 0;
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(stream, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
                SharedStringTable sst = sstpart.SharedStringTable;

                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                Worksheet sheet = worksheetPart.Worksheet;

                var cells = sheet.Descendants<Cell>();
                var rows = sheet.Descendants<Row>();

                log.LogInformation(string.Format("Row count = {0}", rows.LongCount()));
                log.LogInformation(string.Format("Cell count = {0}", cells.LongCount()));

                // One way: go through each cell in the sheet
                foreach (Cell cell in cells)
                {
                    if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
                    {
                        int ssid = int.Parse(cell.CellValue.Text);
                        string str = sst.ChildElements[ssid].InnerText;
                        log.LogInformation(string.Format("Shared string {0}: {1}", ssid, str));
                    }
                    else if (cell.CellValue != null)
                    {
                        log.LogInformation(string.Format("Cell contents: {0}", cell.CellValue.Text));
                    }
                }
            }

            return "Success";
        }
    }
}
