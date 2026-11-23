#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.Calculations;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.Calculations {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelCalculateWorkbookController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("formulas/calculate")]
        public async Task<IActionResult> Calculate(IFormFile file, [FromForm] ExcelCalculateWorkbookModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            var outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var outputStream = CalculateWorkbook(
                inputStream,
                model.DocumentCalculationMode,
                model.IsPdfOutput,
                outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CalculateWorkbook(Stream inputStream, CalculationMode? documentCalculationMode, bool exportToPdf, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            workbook.Calculate();

            if(documentCalculationMode.HasValue)
                workbook.CalculationOptions.Mode = documentCalculationMode.Value;

            var outputStream = new MemoryStream();

            if(exportToPdf)
                workbook.ExportToPdf(outputStream);
            else
                workbook.Save(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
