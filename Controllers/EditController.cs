using IronXL;
using Microsoft.AspNetCore.Mvc;

namespace Ext_IronXL_Project.Controllers
{
    public class EditController : Controller
    {
        [HttpGet]
        public IActionResult InsertRowsColumns()
        {
            return View();
        }

        [HttpPost]
        public IActionResult InsertRowsColumns(IFormFile excelFile, int rowNumber, int columnNumber)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            try
            {
                if (excelFile == null || excelFile.Length == 0)
                {
                    TempData["Message"] = "No file uploaded.";
                    return RedirectToAction("InsertRowsColumns");
                }

                string filesDir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
                if (!Directory.Exists(filesDir)) Directory.CreateDirectory(filesDir);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(filesDir, fileName);
                string outputFilePath = Path.Combine(filesDir, "modified_" + fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                if (!System.IO.File.Exists(filePath))
                {
                    TempData["Message"] = "File not saved.";
                    return RedirectToAction("InsertRowsColumns");
                }

                WorkBook workBook = WorkBook.Load(filePath);
                if (workBook == null)
                {
                    TempData["Message"] = "WorkBook is null after loading.";
                    return RedirectToAction("InsertRowsColumns");
                }

                WorkSheet workSheet = workBook.WorkSheets.FirstOrDefault();
                if (workSheet == null)
                {
                    TempData["Message"] = "No worksheet found in workbook.";
                    return RedirectToAction("InsertRowsColumns");
                }

                // Safely insert row and column (index = rowNumber-1)
                int insertRowIndex = Math.Max(0, rowNumber - 1);
                int insertColIndex = Math.Max(0, columnNumber - 1);

                workSheet.InsertRow(insertRowIndex);
                workSheet.InsertColumn(insertColIndex);

                workBook.SaveAs(outputFilePath);

                TempData["Message"] = "Success! Row and column inserted.";
                TempData["DownloadPath"] = "/files/modified_" + fileName;
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Exception: " + ex.Message;
            }

            return RedirectToAction("InsertRowsColumns");
        }

        [HttpGet]
        public IActionResult GroupUngroup()
        {
            return View();
        }

        [HttpPost]
        public IActionResult GroupUngroup(IFormFile excelFile, int groupRowFrom, int groupRowTo, int ungroupRowFrom, int ungroupRowTo,
                                          string groupColFrom, string groupColTo, string ungroupColFrom, string ungroupColTo)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Message"] = "Please upload a valid Excel file.";
                return RedirectToAction("GroupUngroup");
            }

            try
            {
                string dirPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(dirPath, fileName);
                string outputFile = Path.Combine(dirPath, "grouped_" + fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet workSheet = workBook?.DefaultWorkSheet ?? throw new Exception("Worksheet not found.");

                int rowCount = workSheet.RowCount;
                int colCount = workSheet.ColumnCount;

                
                if (groupRowFrom > 0 && groupRowTo >= groupRowFrom && groupRowTo <= rowCount)
                {
                    workSheet.GroupRows(groupRowFrom - 1, groupRowTo - 1);
                }

                
                if (ungroupRowFrom > 0 && ungroupRowTo >= ungroupRowFrom && ungroupRowTo <= rowCount)
                {
                    workSheet.UngroupRows(ungroupRowFrom - 1, ungroupRowTo - 1);
                }

                
                if (!string.IsNullOrEmpty(groupColFrom) && !string.IsNullOrEmpty(groupColTo))
                {
                    workSheet.GroupColumns(groupColFrom.Trim().ToUpper(), groupColTo.Trim().ToUpper());
                }

                
                if (!string.IsNullOrEmpty(ungroupColFrom) && !string.IsNullOrEmpty(ungroupColTo))
                {
                    workSheet.UngroupColumn(ungroupColFrom.Trim().ToUpper(), ungroupColTo.Trim().ToUpper());
                }

                workBook.SaveAs(outputFile);

                TempData["Message"] = "Grouping and ungrouping completed successfully.";
                TempData["DownloadPath"] = "/files/grouped_" + fileName;
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }

            return RedirectToAction("GroupUngroup");
        }

        [HttpGet]
        public IActionResult RepeatRowsColumns()
        {
            return View();
        }


        [HttpPost]
        public IActionResult RepeatRowsColumns(IFormFile excelFile, int repeatRowFrom, int repeatRowTo,
                                       int repeatColFrom, int repeatColTo, int columnBreakAfter)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Message"] = "Please upload a valid Excel file.";
                return RedirectToAction("RepeatRowsColumns");
            }

            try
            {
                string dirPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(dirPath, fileName);
                string outputPath = Path.Combine(dirPath, "repeat_" + fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet workSheet = workBook?.DefaultWorkSheet ?? throw new Exception("Worksheet not found.");

                int rowCount = workSheet.RowCount;
                int colCount = workSheet.ColumnCount;

                // Validate and apply repeating rows
                if (repeatRowFrom > 0 && repeatRowTo >= repeatRowFrom && repeatRowTo <= rowCount)
                    workSheet.SetRepeatingRows(repeatRowFrom - 1, repeatRowTo - 1);

                // Validate and apply repeating columns
                if (repeatColFrom > 0 && repeatColTo >= repeatColFrom && repeatColTo <= colCount)
                    workSheet.SetRepeatingColumns(repeatColFrom - 1, repeatColTo - 1);

                // Validate and apply column break
                if (columnBreakAfter > 0 && columnBreakAfter <= colCount)
                    workSheet.SetColumnBreak(columnBreakAfter - 1);

                workBook.SaveAs(outputPath);

                TempData["Message"] = "Repeat settings applied successfully.";
                TempData["DownloadPath"] = "/files/repeat_" + fileName;
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }

            return RedirectToAction("RepeatRowsColumns");
        }

        [HttpGet]
        public IActionResult CopyWorksheet()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CopyWorksheet(IFormFile excelFile, string sameBookSheetName, string newBookSheetName)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Message"] = "Please upload a valid Excel file.";
                return RedirectToAction("CopyWorksheet");
            }

            try
            {
                string dir = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "files");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

                string originalFileName = Path.GetFileName(excelFile.FileName);
                string originalFilePath = Path.Combine(dir, originalFileName);

                // Save uploaded file
                using (var stream = new FileStream(originalFilePath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                // Load first workbook
                WorkBook firstBook = WorkBook.Load(originalFilePath);
                WorkSheet originalSheet = firstBook.DefaultWorkSheet ?? throw new Exception("No worksheet found.");

                // Duplicate within same workbook
                originalSheet.CopySheet(sameBookSheetName);

                string sameBookPath = Path.Combine(dir, "same_" + originalFileName);
                firstBook.SaveAs(sameBookPath);

                // Copy to new workbook
                WorkBook secondBook = WorkBook.Create();
                originalSheet.CopyTo(secondBook, newBookSheetName);
                string newBookPath = Path.Combine(dir, "newbook_" + originalFileName);
                secondBook.SaveAs(newBookPath);

                TempData["Message"] = "Worksheet copied successfully!";
                TempData["SameBookPath"] = "/files/same_" + originalFileName;
                TempData["NewBookPath"] = "/files/newbook_" + originalFileName;
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }

            return RedirectToAction("CopyWorksheet");
        }

    }




}
