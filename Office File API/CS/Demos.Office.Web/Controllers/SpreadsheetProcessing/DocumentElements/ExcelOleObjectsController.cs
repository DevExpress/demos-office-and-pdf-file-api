#region Code
using DevExpress.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;
using BlazorDemo.DemoData;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelOleObjectsController : SpreadsheetControllerBase {
        [HttpPost("ole-objects/generate")]
        public async Task<IActionResult> Generate([FromForm] IFormFile file, [FromForm] ExcelOleObjectsModel model) {
            // Read file stream
            using var oleObjectStream = new MemoryStream();
            await file.CopyToAsync(oleObjectStream);
            oleObjectStream.Position = 0;

            Stream imageStream = null;

            if(model.DisplayAsIcon) {
                using var imageTemplateStream = new FileStream(DemoDocumentPaths.WordOleExcelIcon, FileMode.Open, FileAccess.Read);
                imageStream = new MemoryStream();
                await imageTemplateStream.CopyToAsync(imageStream);
                imageStream.Position = 0;
            }

            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate a workbook with an embedded OLE object
            var outputStream = GenerateWorkbookWithOleObject(
                oleObjectStream,
                OleObjectType.ExcelWorksheet,
                outputFormat,
                imageStream);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream GenerateWorkbookWithOleObject(Stream oleObjectStream, OleObjectType oleObjectType, DocumentFormat outputFormat, Stream oleIconStream = null) {
            using var workbook = new Workbook();
            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Ole Object";

            string oleObjectRangeReference = "B2:N34";
            SpreadsheetImageSource oleObjectIcon = null;

            // Create SpreadsheetImageSource if an icon stream is specified
            if(oleIconStream != null) {
                oleObjectRangeReference = "B2:C6";
                oleObjectIcon = SpreadsheetImageSource.FromStream(oleIconStream);
            }

            var oleObjectRange = worksheet.Range[oleObjectRangeReference];

            // Create an embedded OLE object
            var oleObjectEmbedded = worksheet.OleObjects.AddEmbeddedOleObject(oleObjectRange, oleObjectStream, oleObjectType, oleObjectIcon);

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        #endregion Code
    }
}
