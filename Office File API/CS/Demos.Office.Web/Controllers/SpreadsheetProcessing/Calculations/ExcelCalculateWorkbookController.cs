#region Code
using DevExpress.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Calculations {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelCalculateWorkbookController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("formulas/calculate")]
        public async Task<IActionResult> Calculate(IFormFile file, [FromForm] ExcelCalculateWorkbookModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Calculate workbook and convert it to the target format
            var outputStream = CalculateWorkbook(
                inputStream,
                model.DocumentCalculationMode,
                model.IsPdfOutput,
                outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CalculateWorkbook(Stream inputStream, CalculationMode? documentCalculationMode, bool exportToPdf, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Load the document
            workbook.LoadDocument(inputStream);

            // You can use different methods to force recalculation depending on your usage scenario.

            // The "Calculate" method recalculates values in cells marked for calculation and "CalculateAlways" cells:
            // -The cell contains a volatile function (IFunction.Volatile)
            // -The cell references another "CalculateAlways" cell
            // -The cell contains a circular reference
            workbook.Calculate();

            // The "CalculateFull" method forces recalculation of all spreadsheet cells regardless of whether the cell is marked for calculation.
            //workbook.CalculateFull();

            // The "CalculateFullRebuild" method forces a full recalculation of all spreadsheet cells and rebuilds the dependency tree (by default). 
            // This method behavior depends on the DocumentOptions.CalculationEngineType value.
            // When DocumentOptions.CalculationEngineType is set to CalculationEngineType.Recursive CalculateFullRebuild and CalculateFull are equivalent.
            //workbook.CalculateFullRebuild();

            // Change the document calculation mode stored in the resulting document.
            if(documentCalculationMode.HasValue)
                workbook.DocumentSettings.Calculation.Mode = documentCalculationMode.Value;

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            if(exportToPdf)
                workbook.ExportToPdf(outputStream);
            else
                workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
