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
        public IActionResult ExportToExcel(string tableName, string connectionString)
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

                // Load dataset into Excel
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

    }
}
