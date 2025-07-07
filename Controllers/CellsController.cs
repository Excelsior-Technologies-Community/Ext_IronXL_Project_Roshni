using IronXL;
using IronXL.Styles;
using IronXL.Formatting.Enums;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Ext_IronXL_Project.Controllers
{
    public class CellsController : Controller
    {
        [HttpGet]
        public IActionResult CombineRanges()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> CombineRanges(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Error"] = "Please upload a valid Excel file.";
                return RedirectToAction("CombineRanges");
            }

            try
            {
                // Save uploaded file to wwwroot/uploads
                string uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadsFolder))
                    Directory.CreateDirectory(uploadsFolder);

                string filePath = Path.Combine(uploadsFolder, Path.GetFileName(excelFile.FileName));

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await excelFile.CopyToAsync(stream);
                }

                // Load and process Excel file
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet sheet = workBook.WorkSheets.FirstOrDefault();

                if (sheet == null)
                {
                    TempData["Error"] = "No worksheet found in the Excel file.";
                    return RedirectToAction("CombineRanges");
                }

                // Avoid 'Range' ambiguity
                IronXL.Range range1 = sheet["A1:A10"];
                IronXL.Range range2 = sheet["B1:B10"];
                IronXL.Range combinedRange = range1 + range2;

                var values = new List<string>();
                foreach (var cell in combinedRange)
                {
                    values.Add(cell.Text);
                }

                TempData["Success"] = "Ranges combined successfully.";
                ViewBag.RangeValues = values;
                return View();
            }
            catch (Exception ex)
            {
                TempData["Error"] = "An error occurred while processing the Excel file: " + ex.Message;
                return RedirectToAction("CombineRanges");
            }
        }

        [HttpGet]
        public IActionResult StyleExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> StyleExcel(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Error"] = "Please upload a valid Excel file.";
                return RedirectToAction("StyleExcel");
            }

            try
            {
                // Upload location
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(uploadPath, fileName);

                // Save the uploaded file
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await excelFile.CopyToAsync(stream);
                }

                // Load workbook and get worksheet
                WorkBook workbook = WorkBook.Load(filePath);
                WorkSheet sheet = workbook.WorkSheets.First();

                // Apply styles to A1:H10 range
                IronXL.Range range = sheet["A1:H10"];
                var cell = range.First();

                // Style settings
                cell.Style.SetBackgroundColor("#428D65");

                range.Style.Font.Bold = true;
                range.Style.Font.Italic = false;
                range.Style.Font.Underline = FontUnderlineType.SingleAccounting;
                range.Style.Font.Strikeout = false;
                range.Style.Font.FontScript = FontScript.Super;

                range.Style.BottomBorder.Type = BorderType.MediumDashed;
                range.Style.DiagonalBorder.Type = BorderType.Thick;
                range.Style.DiagonalBorderDirection = DiagonalBorderDirection.Forward;
                range.Style.DiagonalBorder.SetColor("#20C96F");

                range.Style.SetBackgroundColor(System.Drawing.Color.Aquamarine);
                range.Style.FillPattern = FillPattern.Diamonds;
                range.Style.VerticalAlignment = VerticalAlignment.Bottom;
                range.Style.Indention = 5;
                range.Style.ShrinkToFit = true;
                range.Style.WrapText = true;

                // Save as styled file
                string styledPath = Path.Combine(uploadPath, "styled_" + fileName);
                workbook.SaveAs(styledPath);

                TempData["Success"] = "Styling applied successfully!";
                ViewBag.DownloadLink = "/uploads/styled_" + fileName;
                return View();
            }
            catch (System.Exception ex)
            {
                TempData["Error"] = "Error processing Excel file: " + ex.Message;
                return RedirectToAction("StyleExcel");
            }
        }


        [HttpGet]
        public IActionResult Analyze()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> Analyze(IFormFile excelFile, string cellRange, string functionType)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0 || string.IsNullOrWhiteSpace(cellRange))
            {
                TempData["Error"] = "Please upload a file and enter a valid cell range.";
                return RedirectToAction("Analyze");
            }

            try
            {
                // Save uploaded file to wwwroot/uploads
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(uploadPath, fileName);

                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await excelFile.CopyToAsync(stream);
                }

                // Load Excel file
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet workSheet = workBook.WorkSheets.First();
                var range = workSheet[cellRange];

                decimal result = 0;
                string label = "";

                // Apply selected function
                switch (functionType)
                {
                    case "Sum":
                        result = range.Sum();
                        label = "Sum";
                        break;
                    case "Average":
                        result = range.Avg();
                        label = "Average";
                        break;
                    case "Max":
                        result = range.Max();
                        label = "Maximum";
                        break;
                    case "Min":
                        result = range.Min();
                        label = "Minimum";
                        break;
                    default:
                        TempData["Error"] = "Invalid function type.";
                        return RedirectToAction("Analyze");
                }

                ViewBag.FunctionLabel = label;
                ViewBag.Result = result;
                ViewBag.Range = cellRange;
                ViewBag.FileName = fileName;

                return View();
            }
            catch (System.Exception ex)
            {
                TempData["Error"] = "Error processing file: " + ex.Message;
                return RedirectToAction("Analyze");
            }
        }

        [HttpGet]
        public IActionResult FormatExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> FormatExcel(IFormFile excelFile, string cellAddress, string formatString)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || string.IsNullOrEmpty(cellAddress) || string.IsNullOrEmpty(formatString))
            {
                TempData["Error"] = "Please upload a file, enter a cell address, and a format string.";
                return RedirectToAction("FormatExcel");
            }

            try
            {
                string uploadPath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadPath))
                    Directory.CreateDirectory(uploadPath);

                string fileName = Path.GetFileName(excelFile.FileName);
                string filePath = Path.Combine(uploadPath, fileName);

                // Save uploaded file
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    await excelFile.CopyToAsync(stream);
                }

                // Load and format Excel file
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet sheet = workBook.WorkSheets.First();

                // Example value to apply format to
                if (formatString.Contains("0")) // number formatting
                    sheet[cellAddress].Value = 123.456;
                else if (formatString.Contains("%"))
                    sheet[cellAddress].Value = 0.78;
                else if (formatString.Contains("yy") || formatString.Contains("d"))
                    sheet[cellAddress].Value = DateTime.Now;

                sheet[cellAddress].FormatString = formatString;

                // Save formatted file
                string formattedFilePath = Path.Combine(uploadPath, "formatted_" + fileName);
                workBook.SaveAs(formattedFilePath);

                ViewBag.FormattedFilePath = "/uploads/formatted_" + fileName;
                ViewBag.CellAddress = cellAddress;
                ViewBag.Format = formatString;

                return View();
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error: " + ex.Message;
                return RedirectToAction("FormatExcel");
            }
        }

        [HttpGet]
        public IActionResult ApplyConditional()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ApplyConditional(IFormFile excelFile, string range, string operatorType, string value1, string value2, string backgroundColor, string fontColor)
        {
            if (excelFile == null || string.IsNullOrEmpty(range) || string.IsNullOrEmpty(operatorType) || string.IsNullOrEmpty(value1))
            {
                TempData["Error"] = "Please provide all required inputs.";
                return RedirectToAction("ApplyConditional");
            }

            try
            {
                // Save uploaded file
                string uploadsFolder = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
                if (!Directory.Exists(uploadsFolder)) Directory.CreateDirectory(uploadsFolder);

                string filePath = Path.Combine(uploadsFolder, excelFile.FileName);
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                // Load Excel file
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet sheet = workBook.DefaultWorkSheet;

                // Parse operator
                ComparisonOperator op = Enum.TryParse(operatorType, out ComparisonOperator parsedOp)
                    ? parsedOp
                    : ComparisonOperator.Equal;

                // Create conditional rule
                var rule = (op == ComparisonOperator.Between)
                    ? sheet.ConditionalFormatting.CreateConditionalFormattingRule(op, value1, value2)
                    : sheet.ConditionalFormatting.CreateConditionalFormattingRule(op, value1);

                // Apply formatting
                if (!string.IsNullOrEmpty(fontColor))
                    rule.FontFormatting.FontColor = fontColor;

                if (!string.IsNullOrEmpty(backgroundColor))
                    rule.PatternFormatting.BackgroundColor = backgroundColor;

                // Apply rule to range
                sheet.ConditionalFormatting.AddConditionalFormatting(range, rule);

                // Save formatted file
                string formattedPath = Path.Combine(uploadsFolder, "formatted_" + excelFile.FileName);
                workBook.SaveAs(formattedPath);

                ViewBag.DownloadLink = "/uploads/formatted_" + excelFile.FileName;
                ViewBag.Range = range;
                ViewBag.Operator = operatorType;

                return View();
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Error: " + ex.Message;
                return RedirectToAction("ApplyConditional");
            }
        }

    }
}
