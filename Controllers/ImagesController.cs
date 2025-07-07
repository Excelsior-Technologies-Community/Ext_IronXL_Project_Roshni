using IronXL;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.ComponentModel;
using IronSoftware.Drawing;

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

            
            var workbook = WorkBook.Load(excelPath);
            var worksheet = workbook.DefaultWorkSheet;

            
            worksheet.InsertImage(imagePath, 1, 1, 3, 3);

            
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
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Error"] = "Please upload a valid Excel file.";
                return RedirectToAction("RemoveImages");
            }

            // Save uploaded Excel file
            string folderPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
            Directory.CreateDirectory(folderPath);
            string uploadedPath = Path.Combine(folderPath, Path.GetFileName(excelFile.FileName));

            using (var stream = new FileStream(uploadedPath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            // Load and process the Excel file
            WorkBook workbook = WorkBook.Load(uploadedPath);
            WorkSheet worksheet = workbook.DefaultWorkSheet;

            // Create a copy of the image list to avoid modifying the collection during iteration
            var images = worksheet.Images.ToList();

            foreach (var img in images)
            {
                worksheet.RemoveImage(img.Id); // ✅ Safe removal
            }

            // Save updated file
            string updatedFileName = "updated_" + Path.GetFileName(excelFile.FileName);
            string updatedFilePath = Path.Combine(folderPath, updatedFileName);
            workbook.SaveAs(updatedFilePath);

            TempData["DownloadLink"] = "/files/" + updatedFileName;
            TempData["Message"] = "Images removed successfully!";
            return RedirectToAction("RemoveImages");
        }

        [HttpGet]
        public IActionResult ExtractImages()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ExtractImagesFromExcel(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Error"] = "Please upload a valid Excel file.";
                return RedirectToAction("ExtractImages");
            }

            // Create directory
            string wwwRoot = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot");
            string filesPath = Path.Combine(wwwRoot, "files");
            string imagesPath = Path.Combine(wwwRoot, "extracted_images");

            Directory.CreateDirectory(filesPath);
            Directory.CreateDirectory(imagesPath);

            // Save uploaded Excel file
            string excelPath = Path.Combine(filesPath, Path.GetFileName(excelFile.FileName));
            using (var stream = new FileStream(excelPath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            // Load the Excel workbook
            WorkBook workBook = WorkBook.Load(excelPath);
            WorkSheet worksheet = workBook.DefaultWorkSheet;

            // Extract and save images
            var images = worksheet.Images.ToList();
            foreach (var image in images)
            {
                AnyBitmap bitmap = image.ToAnyBitmap();
                string fileName = $"image_{image.Id}.png";
                string filePath = Path.Combine(imagesPath, fileName);
                bitmap.SaveAs(filePath);
            }

            TempData["Message"] = $"{images.Count} image(s) extracted successfully!";
            TempData["ImageFolder"] = "/extracted_images";
            return RedirectToAction("ExtractImages");
        }

    }
}
