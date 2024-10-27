using ExcelAssessment.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Diagnostics;

namespace ExcelAssessment.Controllers
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

        public IActionResult Privacy()
        {
            return View();
        }
        [HttpPost]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("No file uploaded.");

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads", file.FileName);
            using (var stream = new FileStream(filePath, FileMode.Create))
            {
                await file.CopyToAsync(stream);
            }

            return RedirectToAction("ProcessExcel", new { fileName = file.FileName });
        }
        public IActionResult ProcessExcel(string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads", fileName);
            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[0];

            var processedCell = worksheet.Cells["F18"].Value;
            if (processedCell != null && processedCell.ToString() == "Total Sum")
            {
                // If processed, read the existing sum from the designated cell.
                int sumRow1 = worksheet.Dimension.End.Row;
                var totalAfterTaxSum = worksheet.Cells[sumRow1, GetColumnIndex(worksheet, "Total Value After Taxing")].GetValue<double>();

                ViewData["TotalAfterTaxSum"] = totalAfterTaxSum;
                ViewData["FileName"] = fileName;
                return View("Result");
            }

            
            int totalValueAfterTaxColumnIndex = GetColumnIndex(worksheet, "Total Value After Taxing");
            int taxingValueColumnIndex = GetColumnIndex(worksheet, "Taxing Value");
            if (totalValueAfterTaxColumnIndex == -1 || taxingValueColumnIndex == -1)
            {
                return BadRequest("Required columns 'Total Value After Taxing' or 'Taxing Value' not found.");
            }

            int totalValueBeforeTaxColumnIndex = GetColumnIndex(worksheet, "Total Value before Taxing");
            if (totalValueBeforeTaxColumnIndex == -1)
            {
                totalValueBeforeTaxColumnIndex = worksheet.Dimension.End.Column + 1;
                worksheet.Cells[1, totalValueBeforeTaxColumnIndex].Value = "Total Value before Taxing";
                worksheet.Cells[1, totalValueBeforeTaxColumnIndex].Style.Font.Bold = true;
                worksheet.Cells[1, totalValueBeforeTaxColumnIndex].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[1, totalValueBeforeTaxColumnIndex].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
            }

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                var cellValue = worksheet.Cells[row, totalValueBeforeTaxColumnIndex].Value;
                if (cellValue == null || cellValue.ToString() == "")
                {
                    var valueAfterTax = worksheet.Cells[row, totalValueAfterTaxColumnIndex].GetValue<double>();
                    var taxValue = worksheet.Cells[row, taxingValueColumnIndex].GetValue<double>();

                    var totalBeforeTax = valueAfterTax - taxValue;
                    worksheet.Cells[row, totalValueBeforeTaxColumnIndex].Value = totalBeforeTax;
                    worksheet.Cells[row, totalValueBeforeTaxColumnIndex].Style.Numberformat.Format = "#,##0.00";
                }
            }

            
            int sumRow = worksheet.Dimension.End.Row + 1;
            if (worksheet.Cells[sumRow - 1, 6].Value?.ToString() != "Total Sum")
            {
                worksheet.Cells[sumRow, 6].Value = "Total Sum";
                worksheet.Cells[sumRow, 6].Style.Font.Bold = true;

                var columnLetter = GetColumnLetter(totalValueAfterTaxColumnIndex);
                worksheet.Cells[sumRow, totalValueAfterTaxColumnIndex].Formula = $"SUM({columnLetter}2:{columnLetter}{sumRow - 1})";
                worksheet.Cells[sumRow, totalValueAfterTaxColumnIndex].Style.Numberformat.Format = "#,##0.00";

                var beforeTaxColumnLetter = GetColumnLetter(totalValueBeforeTaxColumnIndex);
                worksheet.Cells[sumRow, totalValueBeforeTaxColumnIndex].Formula = $"SUM({beforeTaxColumnLetter}2:{beforeTaxColumnLetter}{sumRow - 1})";
                worksheet.Cells[sumRow, totalValueBeforeTaxColumnIndex].Style.Numberformat.Format = "#,##0.00";

                worksheet.Calculate(); 
            }

            

            using (var range = worksheet.Cells[1, 1, sumRow, totalValueBeforeTaxColumnIndex])
            {
                range.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                range.AutoFitColumns();
            }

            package.Save();

            
            ViewData["TotalAfterTaxSum"] = worksheet.Cells[sumRow, totalValueAfterTaxColumnIndex].GetValue<double>();
            ViewData["FileName"] = fileName;
            return View("Result");
        }

        private string GetColumnLetter(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }
        private int GetColumnIndex(ExcelWorksheet worksheet, string headerName)
        {
            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[1, col].Value?.ToString().Trim() == headerName)
                {
                    return col;
                }
            }
            return -1; 
        }
       

        public IActionResult DownloadModifiedSheet(string fileName)
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot/uploads", fileName);
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }


        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
