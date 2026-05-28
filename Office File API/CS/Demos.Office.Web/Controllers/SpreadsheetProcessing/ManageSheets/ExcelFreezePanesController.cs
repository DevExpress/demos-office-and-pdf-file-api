#region Code
using DevExpress.Spreadsheet;
#endregion Code

using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelFreezePanesController : SpreadsheetControllerBase {
        [HttpPost("freeze-panes/freeze")]
        public async Task<IActionResult> Freeze(IFormFile file, [FromForm] ExcelFreezePanesModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            DocumentFormat outputFormatConverted = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            if(model.RowOffset == 0 && model.ColumnOffset == 0) {
                return BadRequest(new { Success = false, Message = "Both RowOffset and ColumnOffset are set to 0" });
            }
            // Freeze Panes on the first worksheet
            var documentStream = await FreezePanes(inputStream, model.RowOffset, model.ColumnOffset, outputFormatConverted);

            // Create Output File
            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        async Task<Stream> FreezePanes(Stream inputStream, int rowOffset, int columnOffset, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            var firstWorksheet = workbook.Worksheets.First();

            // Freeze Panes
            if(rowOffset > 0 && columnOffset > 0) {
                // Freeze both rows and columns
                firstWorksheet.FreezePanes(rowOffset - 1, columnOffset - 1);
            } else {
                if(rowOffset == 0) {
                    // Freeze columns only
                    firstWorksheet.FreezeColumns(columnOffset - 1);
                }
                if(columnOffset == 0) {
                    // Freeze rows only
                    firstWorksheet.FreezeRows(rowOffset - 1);
                }
            }
            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the resulting document
            await workbook.SaveDocumentAsync(outputStream, outputFormat);
            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
