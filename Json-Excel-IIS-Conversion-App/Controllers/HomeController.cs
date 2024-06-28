using Json_Excel_IIS_Conversion_App.Models;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Text;

namespace Json_Excel_IIS_Conversion_App.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult Index(IFormFile formFile)
        {
            if (formFile == null || formFile.Length == 0)
            {
                TempData["error"] = "Please choose one valid JSON file for Conversion";
                return View();
            }

            string fileName = Path.GetFileName(formFile.FileName);
            string contentType = formFile.ContentType;
            byte[] fileContents;
            try
            {
                
                using (var reader = new StreamReader(formFile.OpenReadStream(), Encoding.UTF8))
                {
                    string jsonData = reader.ReadToEnd();
                    List<Sites> sites = JsonConvert.DeserializeObject<List<Sites>>(jsonData);
                    //ConvertToExcel(sites);
                    
                    //#region Excel Convert using EPPlus
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                    using (var excelPackage = new ExcelPackage())
                    {


                        foreach (var config in sites)
                        {
                            // Add a new worksheet to the empty workbook
                            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(config.SiteName);

                            // Define header columns
                            worksheet.Cells[1, 1].Value = "Issue Type";
                            worksheet.Cells[1, 2].Value = "Issue Parameter";
                            worksheet.Cells[1, 3].Value = "Parameter";
                            worksheet.Cells[1, 4].Value = "Recommendations";

                            // Populate data
                            int row = 2;

                            foreach (var failedchecks in config.FailedChecks)
                            {
                                //Issue Type

                                worksheet.Cells[row, 1].Value = "FailedChecks";
                                //Issue Parameter
                                worksheet.Cells[row, 2].Value = failedchecks.IssueId;

                                //Parameter
                                worksheet.Cells[row, 3].Value = failedchecks.Details;
                                worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                                worksheet.Cells[row, 3].Style.WrapText = true;

                                //Recommendations
                                worksheet.Cells[row, 4].Value = failedchecks.Recommendation;
                                worksheet.Cells[row, 4].Style.WrapText = true;

                                row++;
                            }

                            foreach (var warningChecks in config.WarningChecks)
                            {
                                //Issue Type

                                worksheet.Cells[row, 1].Value = "WarningChecks";
                                //Issue Parameter
                                worksheet.Cells[row, 2].Value = warningChecks.IssueId;

                                //Parameter
                                worksheet.Cells[row, 3].Value = warningChecks.Details;
                                worksheet.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                worksheet.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightYellow);
                                worksheet.Cells[row, 3].Style.WrapText = true;

                                //Recommendations
                                worksheet.Cells[row, 4].Value = warningChecks.Recommendation;
                                worksheet.Cells[row, 4].Style.WrapText = true;

                                row++;
                            }

                            // Auto-fit columns for better aesthetics
                            worksheet.Cells.AutoFitColumns(0);
                            
                        }

                        fileContents = excelPackage.GetAsByteArray();
                        return File(fileContents, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Path.GetFileNameWithoutExtension(fileName) + ".xlsx");
                    }

                }
            }
            catch
            {
                TempData["error"] = "Error converting json data to excel.";
                return BadRequest(new
                {
                    message = "Error converting json data to excel."
                });
            }           
        }

        [HttpPost]
        public IActionResult Clear()
        {
            TempData["success"] = "Page reset successful.";
            return RedirectToAction("Index");
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
