using Ext_IronXL_Project.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using IronXL;
using System.IO;

namespace Ext_IronXL_Project.Controllers
{
    public class HomeController : Controller
    {
        private readonly IWebHostEnvironment _env;

        public HomeController(IWebHostEnvironment webHostEnvironment)
        {
            _env = webHostEnvironment;
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult CreateExcel()
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            WorkBook workBook = WorkBook.Create();

            // Create a blank WorkSheet
            WorkSheet workSheet = workBook.CreateWorkSheet("new_sheet");

            // Add data and styles
            workSheet["A1"].Value = "Hello World";
            workSheet["A1"].Style.WrapText = true;

            workSheet["A2"].BoolValue = true;
            workSheet["A2"].Style.BottomBorder.Type = IronXL.Styles.BorderType.Double;

            // Save to a temp file on disk
            string tempFilePath = Path.GetTempFileName() + ".xlsx";
            workBook.SaveAs(tempFilePath);

            // Read file into memory stream
            byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
            System.IO.File.Delete(tempFilePath); // Clean up temp file

            // Return file
            return File(fileBytes,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "sample.xlsx");
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
