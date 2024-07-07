using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace TaxesTask.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        [HttpPost]
        public IActionResult ProcessExcel(IFormFile file)
        {
            if (file == null || !Path.GetExtension(file.FileName).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest("Invalid file format. Please upload an .xlsx file.");
            }

            try
            {
                using (var workbook = new XLWorkbook(file.OpenReadStream()))
                {
                    var worksheet = workbook.Worksheet(1);

                    if (worksheet.Column(7).IsEmpty() || worksheet.Column(8).IsEmpty())
                    {
                        return BadRequest("Excel file does not have data in columns 7 and 8.");
                    }

                    if (worksheet.Cell(1, 9).IsEmpty())
                    {
                        worksheet.Cell(1, 9).Value = "Total Value Before Taxing";
                    }

                    var lastRow = worksheet.LastRowUsed().RowNumber();
                    for (int row = 2; row <= lastRow; row++)
                    {
                        decimal value7, value8;
                        if (decimal.TryParse(worksheet.Cell(row, 7).GetValue<string>(), out value7) &&
                            decimal.TryParse(worksheet.Cell(row, 8).GetValue<string>(), out value8))
                        {
                            worksheet.Cell(row, 9).Value = value7 - value8;
                        }
                        else
                        {
                            worksheet.Cell(row, 9).Value = "Invalid data";
                        }
                    }

                    var stream = new MemoryStream();
                    workbook.SaveAs(stream);
                    stream.Seek(0, SeekOrigin.Begin);

                    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "modified_taxes.xlsx");
                }
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }
    }
}
