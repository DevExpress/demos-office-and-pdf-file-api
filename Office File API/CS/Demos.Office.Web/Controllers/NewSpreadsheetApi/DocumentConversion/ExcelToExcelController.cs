#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentConversion;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentConversion {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelToExcelController : SpreadsheetControllerBase {
        [HttpPost("excel-to-excel/convert")]
        public async Task<IActionResult> ConvertExcelToExcel(IFormFile file, [FromForm] ExcelToExcelConversionModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);
            var convertedStream = ConvertDocument(inputStream, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ConvertDocument(Stream inputStream, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
