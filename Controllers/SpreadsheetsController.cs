using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using IronXL;
using System.Data;
using System.Data.SqlClient;
using System.IO;

namespace Ext_IronXL_Project.Controllers
{
    public class SpreadsheetsController : Controller
    {
        [HttpGet]
        public IActionResult ReadExcel()
        {
            return View();
        }
        [HttpPost]
        public IActionResult ReadExcels(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Error = "Please upload a valid Excel file.";
                return View("ReadExcel");
            }

            List<string> cellData = new List<string>();
            decimal sum = 0;
            decimal max = 0;

            using (var stream = new MemoryStream())
            {
                excelFile.CopyTo(stream);
                stream.Position = 0;

                // Load the workbook
                WorkBook workBook = WorkBook.Load(stream);
                WorkSheet workSheet = workBook.DefaultWorkSheet;

                // Read range A2:A10
                var range = workSheet["A2:A10"];

                foreach (var cell in range)
                {
                    cellData.Add($"Cell {cell.AddressString} has value '{cell.Text}'");
                }

                sum = range.Sum();
                max = range.Max(c => c.DecimalValue);
            }

            ViewBag.Cells = cellData;
            ViewBag.Sum = sum;
            ViewBag.Max = max;

            return View("ReadExcel");
        }

        [HttpGet]
        public IActionResult SqlToExcel()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExportSqlToExcel(string tableName, string connectionString)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(connectionString))
            {
                ViewBag.Error = "Please enter both table name and connection string.";
                return View("SqlToExcel");
            }

            string sql = $"SELECT * FROM [{tableName}]";
            DataSet ds = new DataSet();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    SqlDataAdapter adapter = new SqlDataAdapter(sql, connection);
                    adapter.Fill(ds);
                }

                if (ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                {
                    ViewBag.Error = "No data found in the table.";
                    return View("SqlToExcel");
                }

                
                WorkBook workBook = WorkBook.Load(ds);
                workBook.DefaultWorkSheet.Name = tableName;

                byte[] fileBytes = workBook.ToByteArray();

                return File(fileBytes,
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            $"{tableName}_Data.xlsx");


            }
            catch (SqlException ex)
            {
                ViewBag.Error = $"SQL Error: {ex.Message}";
                return View("SqlToExcel");
            }
            catch (System.Exception ex)
            {
                ViewBag.Error = $"Error: {ex.Message}";
                return View("SqlToExcel");
            }
        }

        [HttpGet]
        public IActionResult ExportExcel()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExportExcelFiles()
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            WorkBook workBook = WorkBook.Create();

            
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

            
            workSheet["A1"].Value = "Hello World";
            workSheet["A2"].Style.BottomBorder.SetColor("#ff6600");

            byte[] fileBytes = workBook.ToByteArray();

            
            return File(fileBytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "sample.xlsx");
        }

        public IActionResult ConvertFile()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ConvertFiles(IFormFile excelFile, string format)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0 || string.IsNullOrEmpty(format))
            {
                TempData["Error"] = "Please upload a file and select a format.";
                return RedirectToAction("ConvertFile");
            }

            // Load the uploaded file into IronXL
            using (var stream = new MemoryStream())
            {
                excelFile.CopyTo(stream);
                stream.Position = 0;

                WorkBook workBook = WorkBook.Load(stream);
                byte[] resultBytes;
                string contentType;
                string fileName;

                switch (format.ToLower())
                {
                    case "xlsx":
                        resultBytes = workBook.ToByteArray();
                        contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                        fileName = "converted.xlsx";
                        break;

                    case "csv":
                        string csvPath = Path.GetTempFileName();
                        workBook.SaveAsCsv(csvPath);
                        resultBytes = System.IO.File.ReadAllBytes(csvPath);
                        System.IO.File.Delete(csvPath);
                        contentType = "text/csv";
                        fileName = "converted.csv";
                        break;

                    case "json":
                        string jsonPath = Path.GetTempFileName();
                        workBook.SaveAsJson(jsonPath);
                        resultBytes = System.IO.File.ReadAllBytes(jsonPath);
                        System.IO.File.Delete(jsonPath);
                        contentType = "application/json";
                        fileName = "converted.json";
                        break;

                    case "xml":
                        string xmlPath = Path.GetTempFileName();
                        workBook.SaveAsXml(xmlPath);
                        resultBytes = System.IO.File.ReadAllBytes(xmlPath);
                        System.IO.File.Delete(xmlPath);
                        contentType = "application/xml";
                        fileName = "converted.xml";
                        break;

                    case "html":
                        string htmlContent = workBook.ExportToHtmlString();
                        resultBytes = System.Text.Encoding.UTF8.GetBytes(htmlContent);
                        contentType = "text/html";
                        fileName = "converted.html";
                        break;

                    default:
                        TempData["Error"] = "Unsupported format.";
                        return RedirectToAction("ConvertFile");
                }

                return File(resultBytes, contentType, fileName);
            }
        }
    }
}
