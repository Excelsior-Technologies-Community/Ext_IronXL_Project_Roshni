using IronXL;
using Microsoft.AspNetCore.Mvc;

namespace Ext_IronXL_Project.Controllers
{
    public class ImagesController : Controller
    {
        [HttpGet]
        public IActionResult AddImageToExcel()
        {
            return View();
        }

        [HttpPost]
        public IActionResult AddImageToExcel(IFormFile excelFile, IFormFile imageFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || imageFile == null)
            {
                TempData["Error"] = "Both files are required.";
                return View();
            }

            string folder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
            Directory.CreateDirectory(folder);

            string excelPath = Path.Combine(folder, Path.GetFileName(excelFile.FileName));
            string imagePath = Path.Combine(folder, Path.GetFileName(imageFile.FileName));

            using (var fs = new FileStream(excelPath, FileMode.Create))
                excelFile.CopyTo(fs);

            using (var fs = new FileStream(imagePath, FileMode.Create))
                imageFile.CopyTo(fs);

            // Load Excel and insert image
            var workbook = WorkBook.Load(excelPath);
            var worksheet = workbook.DefaultWorkSheet;

            // Insert image from file at position (Row 2, Col 2) to (Row 4, Col 4)
            worksheet.InsertImage(imagePath, 1, 1, 3, 3);

            // Save updated file
            string updatedPath = Path.Combine(folder, "updated_" + Path.GetFileName(excelFile.FileName));
            workbook.SaveAs(updatedPath);

            TempData["DownloadLink"] = "/files/" + Path.GetFileName(updatedPath);
            return RedirectToAction("AddImageToExcel");
        }

        [HttpGet]
        public IActionResult RemoveImages()
        {
            return View();
        }

        [HttpPost]
        public IActionResult RemoveImages(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Error"] = "Please upload a valid Excel file.";
                return RedirectToAction("RemoveImages");
            }

            // Save uploaded file to wwwroot/files
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
            Directory.CreateDirectory(folderPath);

            string uploadedPath = Path.Combine(folderPath, Path.GetFileName(excelFile.FileName));
            using (var stream = new FileStream(uploadedPath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            // Load workbook and worksheet
            WorkBook workBook = WorkBook.Load(uploadedPath);
            WorkSheet workSheet = workBook.DefaultWorkSheet;

            // Remove all images using Clear() (most reliable way)
            if (workSheet.Images != null && workSheet.Images.Count > 0)
            {
                workSheet.Images.Clear();
            }

            // Save updated file
            string updatedFileName = "updated_" + Path.GetFileName(excelFile.FileName);
            string updatedFilePath = Path.Combine(folderPath, updatedFileName);
            workBook.SaveAs(updatedFilePath);

            TempData["DownloadLink"] = "/files/" + updatedFileName;
            return RedirectToAction("RemoveImages");
        }
    }
}
