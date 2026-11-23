#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.ManageSheets {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelClearContentsController : SpreadsheetControllerBase {
        [HttpPost("clear-contents/clear")]
        public async Task<IActionResult> Clear(IFormFile file, [FromForm] ExcelClearContentsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentStream = ClearContents(inputStream, model.ClearMode, targetFormat);

            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ClearContents(Stream inputStream, ClearMode clearMode, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            CellRange rangeToClear = firstWorksheet.GetUsedRange();

            ClearContentType contentToClear = clearMode switch {
                ClearMode.CellContent => ClearContentType.Values,
                ClearMode.CellFormatting => ClearContentType.Formats,
                ClearMode.CellHyperlinks => ClearContentType.Hyperlinks,
                ClearMode.CellComments => ClearContentType.Comments,
                _ => ClearContentType.All
            };

            rangeToClear.Clear(contentToClear);

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
