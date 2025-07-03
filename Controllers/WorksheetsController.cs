using Microsoft.AspNetCore.Mvc;
using IronXL;
using IronXL.Drawing.Charts;

namespace Ext_IronXL_Project.Controllers
{
    public class WorksheetsController : Controller
    {
        [HttpGet]
        public IActionResult FormulaEditor()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> FormulaEditor(IFormFile excelFile, string formulas)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please select a valid Excel file.";
                return View();
            }

            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            using (var stream = new FileStream(tempPath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            try
            {
                var workBook = WorkBook.Load(tempPath);
                var workSheet = workBook.DefaultWorkSheet;

                // Process formulas
                if (!string.IsNullOrWhiteSpace(formulas))
                {
                    var lines = formulas.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var line in lines)
                    {
                        var parts = line.Split('=', 2);
                        if (parts.Length == 2)
                        {
                            var cell = parts[0].Trim();
                            var formula = parts[1].Trim();

                            workSheet[cell].Formula = formula;
                        }
                    }
                }

                // Recalculate all formulas
                workBook.EvaluateAll();

                // Save updated file
                workBook.SaveAs(tempPath);

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempPath);
                string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName) + "_Formulas.xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
                return View();
            }
            finally
            {
                if (System.IO.File.Exists(tempPath))
                    System.IO.File.Delete(tempPath);
            }
        }

        [HttpGet]
        public IActionResult SortExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> SortExcel(IFormFile excelFile, string range, string sortByColumn, int? columnIndex)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid Excel file.";
                return View();
            }

            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");

            using (var stream = new FileStream(tempPath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            try
            {
                WorkBook workBook = WorkBook.Load(tempPath);
                WorkSheet workSheet = workBook.DefaultWorkSheet;

                // Sort Range (e.g., A1:D20)
                if (!string.IsNullOrWhiteSpace(range))
                {
                    var selectedRange = workSheet[range];
                    selectedRange.SortAscending(); // Sort range ascending

                    if (!string.IsNullOrWhiteSpace(sortByColumn))
                    {
                        selectedRange.SortByColumn(sortByColumn, SortOrder.Ascending); // e.g. by column "C"
                    }
                }

                // Sort a single column if provided
                if (columnIndex.HasValue)
                {
                    var column = workSheet.GetColumn(columnIndex.Value); // 1 = B
                    column.SortDescending();
                }

                workBook.SaveAs(tempPath);

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempPath);
                string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName) + "_Sorted.xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
                return View();
            }
            finally
            {
                if (System.IO.File.Exists(tempPath))
                    System.IO.File.Delete(tempPath);
            }
        }

        [HttpGet]
        public IActionResult SelectRange()
        {
            return View();
        }

        [HttpPost]
        public IActionResult SelectRange(IFormFile excelFile, string range)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0 || string.IsNullOrEmpty(range))
            {
                ViewBag.Message = "Please upload a valid Excel file and enter a range.";
                return View();
            }

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(excelFile.FileName));
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            List<string> rangeValues = new List<string>();
            List<List<string>> allRows = new List<List<string>>();

            try
            {
                WorkBook workBook = WorkBook.Load(filePath);
                WorkSheet sheet = workBook.WorkSheets.First();

                // Get specified range
                var selectedRange = sheet[range];
                foreach (var cell in selectedRange)
                {
                    rangeValues.Add(cell?.Value?.ToString());
                }

                // Get all rows in worksheet
                foreach (var row in sheet.Rows)
                {
                    List<string> rowValues = new List<string>();
                    foreach (var cell in row)
                    {
                        rowValues.Add(cell?.Value?.ToString() ?? "");
                    }
                    allRows.Add(rowValues);
                }

                return View(Tuple.Create(rangeValues, allRows));
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
                return View();
            }
            finally
            {
                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);
            }
        }

        [HttpGet]
        public IActionResult CreateChart()
        {
            return View();
        }

        [HttpPost]
        public IActionResult CreateChart(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid Excel file.";
                return View();
            }

            try
            {
                string uploadPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + Path.GetExtension(excelFile.FileName));
                using (var stream = new FileStream(uploadPath, FileMode.Create))
                {
                    excelFile.CopyTo(stream);
                }

                WorkBook workBook = WorkBook.Load(uploadPath);
                WorkSheet workSheet = workBook.DefaultWorkSheet;

                // Create chart (Line chart at specific location)
                var chart = workSheet.CreateChart(ChartType.Line, 10, 10, 18, 20);
                var series = chart.AddSeries("B3:B8", "A3:A8"); // Y-axis, X-axis
                series.Title = "Line Chart";

                chart.SetLegendPosition(LegendPosition.Bottom);
                chart.Position.LeftColumnIndex = 2;
                chart.Position.RightColumnIndex = chart.Position.LeftColumnIndex + 3;
                chart.Plot();

                // Save the file
                string resultFile = "CreateLineChart_" + Guid.NewGuid().ToString("N") + ".xlsx";
                string resultPath = Path.Combine(Path.GetTempPath(), resultFile);
                workBook.SaveAs(resultPath);

                ViewBag.Message = "Chart created and saved!";
                ViewBag.DownloadPath = resultFile;
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
            }

            return View();
        }

        [HttpGet]
        public IActionResult DownloadChart(string fileName)
        {
            if (string.IsNullOrEmpty(fileName)) return NotFound();

            string filePath = Path.Combine(Path.GetTempPath(), fileName);
            if (!System.IO.File.Exists(filePath)) return NotFound();

            var memory = new MemoryStream();
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                stream.CopyTo(memory);
            }

            memory.Position = 0;
            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [HttpGet]
        public IActionResult FreezePaneForm()
        {
            return View();
        }

        [HttpPost]
        public IActionResult ApplyFreezePane(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile != null && excelFile.Length > 0)
            {
                using var stream = excelFile.OpenReadStream();
                WorkBook workBook = WorkBook.Load(stream);
                WorkSheet workSheet = workBook.WorkSheets.First();

                // Apply freeze panes
                workSheet.CreateFreezePane(2, 3);         // A-B, 1-3
                workSheet.CreateFreezePane(5, 5, 6, 7);   // A-E, 1-5, scroll to F,G..., row 7+

                // Convert to stream
                using var memoryStream = workBook.ToStream();  // default format is XLSX
                return File(memoryStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "FreezePaneResult.xlsx");
            }

            ViewBag.Error = "Please upload a valid Excel file.";
            return View("FreezePaneForm");
        }

    }
}
