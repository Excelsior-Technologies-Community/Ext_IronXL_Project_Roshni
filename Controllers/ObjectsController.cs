using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using Microsoft.Data.SqlClient;
using System.IO;
using IronXL;
using System.Data;
using System.Data.SqlClient;
using IronXL.Printing;
using System.Threading.Tasks;

namespace Ext_IronXL_Project.Controllers
{
    public class ObjectsController : Controller
    {
        private readonly string _connectionString = "Server=DESKTOP-0OMS0D3\\SQLEXPRESS;Database=db1;Trusted_Connection=True;TrustServerCertificate=True;";
        private readonly IWebHostEnvironment _env;

        public ObjectsController(IWebHostEnvironment env)
        {
            _env = env;
        }
        public IActionResult Import()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Import(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                TempData["Message"] = "No file selected.";
                return RedirectToAction("Import");
            }

            var fileExtension = Path.GetExtension(excelFile.FileName).ToLower();
            var supportedExtensions = new[] { ".csv", ".tsv", ".xls", ".xlt", ".xlsm", ".xlsx", ".xltx" };

            if (!supportedExtensions.Contains(fileExtension))
            {
                TempData["Message"] = "Unsupported file format.";
                return RedirectToAction("Import");
            }

            var uploadsFolder = Path.Combine(Path.GetTempPath(), "UploadedExcels");
            Directory.CreateDirectory(uploadsFolder);

            var uniqueFileName = Guid.NewGuid().ToString() + fileExtension;
            var filePath = Path.Combine(uploadsFolder, uniqueFileName);

            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                excelFile.CopyTo(stream);
            }

            try
            {
                WorkBook workbook = WorkBook.Load(filePath);
                DataSet dataSet = workbook.ToDataSet();

                foreach (DataTable table in dataSet.Tables)
                {
                    foreach (DataRow row in table.Rows)
                    {
                        using (SqlConnection conn = new SqlConnection(_connectionString))
                        {
                            conn.Open();

                            
                            string query = "INSERT INTO ImportedData (Column1, Column2, Column3, Column4, Column5) VALUES (@c1, @c2, @c3, @c4, @c5)";
                            using (SqlCommand cmd = new SqlCommand(query, conn))
                            {
                                cmd.Parameters.AddWithValue("@c1", row.Table.Columns.Count > 0 ? row[0]?.ToString() : DBNull.Value);
                                cmd.Parameters.AddWithValue("@c2", row.Table.Columns.Count > 1 ? row[1]?.ToString() : DBNull.Value);
                                cmd.Parameters.AddWithValue("@c3", row.Table.Columns.Count > 2 ? row[2]?.ToString() : DBNull.Value);
                                cmd.Parameters.AddWithValue("@c4", row.Table.Columns.Count > 3 ? row[3]?.ToString() : DBNull.Value);
                                cmd.Parameters.AddWithValue("@c5", row.Table.Columns.Count > 4 ? row[4]?.ToString() : DBNull.Value);

                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }

                TempData["Message"] = "Data imported successfully!";
            }
            catch (Exception ex)
            {
                TempData["Message"] = "Error: " + ex.Message;
            }
            finally
            {
                if (System.IO.File.Exists(filePath))
                    System.IO.File.Delete(filePath);
            }

            return RedirectToAction("Import");
        }

        public IActionResult DataGrid()
        {
            return View();
        }

        [HttpPost]
        public IActionResult DataGrid(IFormFile excelFile)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please select a valid Excel file.";
                return View();
            }

            DataTable dataTable = new DataTable();

            try
            {
                using (var stream = new MemoryStream())
                {
                    excelFile.CopyTo(stream);
                    stream.Position = 0;

                    
                    WorkBook workBook = WorkBook.Load(stream);
                    WorkSheet workSheet = workBook.DefaultWorkSheet;

                    
                    dataTable = workSheet.ToDataTable(true);
                }
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
                return View();
            }

            return View(dataTable);
        }

        public IActionResult Update()
        {
            return View("Update");
        }

        [HttpPost]
        public IActionResult Update(IFormFile excelFile, string tableName)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0 || string.IsNullOrWhiteSpace(tableName))
            {
                ViewBag.Message = "Please provide a valid Excel file and table name.";
                return View();
            }

            try
            {
                using var stream = excelFile.OpenReadStream();
                var workBook = WorkBook.Load(stream);
                var sheet = workBook.WorkSheets.First();

                var headers = sheet.Rows[0].Select(cell => cell.Text).ToList(); 

                using SqlConnection conn = new SqlConnection(_connectionString);
                conn.Open();

                for (int row = 1; row < sheet.RowCount; row++) 
                {
                    var data = sheet.Rows[row];

                    
                    Dictionary<string, object> rowData = new Dictionary<string, object>();
                    for (int col = 0; col < headers.Count; col++)
                    {
                        string colName = headers[col];
                        object value = data.Columns[col].Value;
                        rowData[colName] = value ?? DBNull.Value;
                    }

                    
                    if (!rowData.ContainsKey("Id"))
                        continue;

                    var idValue = rowData["Id"];

                    
                    string checkSql = $"SELECT COUNT(*) FROM {tableName} WHERE Id = @Id";
                    using var checkCmd = new SqlCommand(checkSql, conn);
                    checkCmd.Parameters.AddWithValue("@Id", idValue);
                    int exists = (int)checkCmd.ExecuteScalar();

                    if (exists > 0)
                    {
                        
                        var setParts = headers.Where(h => h != "Id").Select(h => $"{h} = @{h}");
                        string updateSql = $"UPDATE {tableName} SET {string.Join(", ", setParts)} WHERE Id = @Id";
                        using var updateCmd = new SqlCommand(updateSql, conn);

                        
                        foreach (var kvp in rowData)
                            updateCmd.Parameters.AddWithValue("@" + kvp.Key, kvp.Value ?? DBNull.Value);

                        updateCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        
                        string cols = string.Join(", ", headers);
                        string vals = string.Join(", ", headers.Select(h => "@" + h));
                        string insertSql = $"INSERT INTO {tableName} ({cols}) VALUES ({vals})";
                        using var insertCmd = new SqlCommand(insertSql, conn);

                        
                        foreach (var kvp in rowData)
                            insertCmd.Parameters.AddWithValue("@" + kvp.Key, kvp.Value ?? DBNull.Value);

                        insertCmd.ExecuteNonQuery();
                    }
                }

                ViewBag.Message = "Excel data uploaded and synced with database.";
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
            }

            return View();
        }

        [HttpGet]
        public IActionResult EditMetadata()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> EditMetadata(IFormFile excelFile, string author, string title, string comments, string keywords)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please select a valid Excel file.";
                return View();
            }

            try
            {
                using var stream = new MemoryStream();
                await excelFile.CopyToAsync(stream);
                stream.Position = 0;

                var workBook = WorkBook.Load(stream);

                
                workBook.Metadata.Author = author;
                workBook.Metadata.Title = title;
                workBook.Metadata.Comments = comments;
                workBook.Metadata.Keywords = keywords;

                
                DateTime? created = workBook.Metadata.Created;
                DateTime? printed = workBook.Metadata.LastPrinted;

                
                string outputDir = Path.Combine(_env.WebRootPath, "downloads");
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                string filePath = Path.Combine(outputDir, $"edited_{Path.GetFileName(excelFile.FileName)}");
                workBook.SaveAs(filePath);

                ViewBag.Message = $"Metadata updated successfully. File saved at: /downloads/{Path.GetFileName(filePath)}";
            }
            catch (Exception ex)
            {
                ViewBag.Message = "Error: " + ex.Message;
            }

            return View();
        }

        [HttpGet]
        public IActionResult ProtectExcel()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> ProtectExcel(IFormFile excelFile, string password, string action)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please select a valid Excel file.";
                return View();
            }

            // Ensure correct extension
            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            try
            {
                WorkBook workBook = null;

                if (action == "decrypt")
                {
                    workBook = WorkBook.Load(tempFilePath, password);
                    workBook.Password = null;
                    workBook.SaveAs(tempFilePath);
                    ViewBag.Message = "Workbook decrypted successfully.";
                }
                else
                {
                    workBook = WorkBook.Load(tempFilePath);
                    WorkSheet sheet = workBook.DefaultWorkSheet;

                    switch (action)
                    {
                        case "encrypt":
                            workBook.Encrypt(password);
                            workBook.SaveAs(tempFilePath);
                            ViewBag.Message = "Workbook encrypted successfully.";
                            break;

                        case "protectSheet":
                            sheet.ProtectSheet(password);
                            workBook.SaveAs(tempFilePath);
                            ViewBag.Message = "Worksheet protected successfully.";
                            break;

                        case "unprotectSheet":
                            sheet.UnprotectSheet();
                            workBook.SaveAs(tempFilePath);
                            ViewBag.Message = "Worksheet unprotected successfully.";
                            break;
                    }
                }

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName) + "_protected.xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch
            {
                ViewBag.Message = "Error processing the Excel file. Check if the password is correct.";
                return View();
            }
            finally
            {
                if (System.IO.File.Exists(tempFilePath))
                    System.IO.File.Delete(tempFilePath);
            }
        }

        [HttpGet]
        public IActionResult PrintSetup()
        {
            return View();
        }

        [HttpPost]
        public async Task<IActionResult> PrintSetup(IFormFile excelFile, string headerText, string footerText, double headerMargin, string paperSize, string orientation, bool noColor)
        {
            IronXL.License.LicenseKey = "IRONSUITE.EXCELSIORTECHNOLOGIES.8187-3708CDF60F-B3MFZGNEBBLCW6W5-BXTNM4MPHNMU-OTOEBFCLRDSW-ETTEAFOQHJNT-KBYQIADCMQPR-QVRLMSV47S2Q-GLWMCW-TRH25WCBAZKQUA-SANDBOX-BDIXLR.TRIAL.EXPIRES.27.MAR.2026";

            if (excelFile == null || excelFile.Length == 0)
            {
                ViewBag.Message = "Please upload a valid Excel file.";
                return View();
            }

            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");

            using (var stream = new FileStream(tempFilePath, FileMode.Create))
            {
                await excelFile.CopyToAsync(stream);
            }

            try
            {
                var workBook = WorkBook.Load(tempFilePath);
                var workSheet = workBook.DefaultWorkSheet;

                // Header and Footer
                workSheet.Header.Center = headerText;
                workSheet.Footer.Center = footerText;

                // Header Margin
                workSheet.PrintSetup.HeaderMargin = headerMargin;

                // Paper Size
                if (Enum.TryParse<PaperSize>(paperSize, out var size))
                    workSheet.PrintSetup.PaperSize = size;

                // Orientation
                workSheet.PrintSetup.PrintOrientation = orientation == "Landscape" ? PrintOrientation.Landscape : PrintOrientation.Portrait;

                // Color settings
                workSheet.PrintSetup.NoColor = noColor;

                workBook.SaveAs(tempFilePath);

                byte[] fileBytes = System.IO.File.ReadAllBytes(tempFilePath);
                string fileName = Path.GetFileNameWithoutExtension(excelFile.FileName) + "_PrintSetup.xlsx";
                return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
            catch
            {
                ViewBag.Message = "Error processing the Excel file.";
                return View();
            }
            finally
            {
                if (System.IO.File.Exists(tempFilePath))
                    System.IO.File.Delete(tempFilePath);
            }
        }


    }
}
