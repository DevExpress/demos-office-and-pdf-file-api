#region Code
using DevExpress.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentConversion;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.Spreadsheet.DocumentConversion {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelToExcelController : SpreadsheetControllerBase {
        [HttpPost("excel-to-excel/convert")]
        public async Task<IActionResult> ConvertExcelToExcel(IFormFile file, [FromForm] ExcelToExcelConversionModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);
            // Convert document to the target format
            var convertedStream = ConvertDocument(inputStream, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ConvertDocument(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            // Load the input document
            workbook.LoadDocument(inputStream);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
