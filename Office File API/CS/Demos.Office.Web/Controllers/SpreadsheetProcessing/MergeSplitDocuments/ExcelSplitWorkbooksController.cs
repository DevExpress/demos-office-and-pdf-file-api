#region Code
using DevExpress.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.MergeSplitDocuments;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.MergeSplitDocuments {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelSplitWorkbooksController : SpreadsheetControllerBase {
        [HttpPost("split-workbooks/split")]
        public async Task<IActionResult> Split(IFormFile file, [FromForm] ExcelSplitWorkbooksModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;
            DocumentFormat outputFormatConverted = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Split Workbook
            var documentStreams = SplitWorkbook(model.SplitMode == VisibilityMode.VisibleOnly, inputStream, outputFormatConverted);

            // Create Output File
            if(documentStreams == null || documentStreams.Count == 0)
                return StatusCode(500, new { Success = false, Message = "Failed to split workbook" });

            if(documentStreams.Count == 1) {
                // Single worksheet return directly
                var fileName = "document";
                return CreateOutputFile((Stream)documentStreams[0], model.OutputFormat, fileName);
            } else {
                // Multiple worksheets - create ZIP file
                var zipFileName = $"{model.OutputName}.zip";
                return CreateZipOutputFile(documentStreams, model.OutputFormat, "worksheet", zipFileName);
            }
        }
        #region Code
        IReadOnlyList<Stream> SplitWorkbook(bool splitVisibleOnly, Stream inputStream, DocumentFormat outputFormat) {
            // Load document from stream
            using var sourceWorkbook = new Workbook();

            sourceWorkbook.LoadDocument(inputStream);

            // Split document worksheets to separate documents
            var worksheetsToSplit = new List<Worksheet>();
            if(splitVisibleOnly) {
                // Get visible worksheets only
                worksheetsToSplit = sourceWorkbook.Worksheets.Where(worksheet => worksheet.Visible).ToList();
            } else {
                // Get all worksheets except "Very hidden"
                worksheetsToSplit = sourceWorkbook.Worksheets.Where(worksheet => worksheet.VisibilityType != WorksheetVisibilityType.VeryHidden).ToList();
            }

            // Create separate documents from selected worksheets
            var documentStreams = new List<Stream>();
            foreach(var sourceWorksheet in worksheetsToSplit) {
                using var targetWorkbook = new Workbook();
                var targetWorksheet = targetWorkbook.Worksheets[0];
                targetWorksheet.Name = sourceWorksheet.Name;
                targetWorksheet.CopyFrom(sourceWorksheet);

                // Ensure the copied worksheet is visible in the output document
                targetWorksheet.VisibilityType = WorksheetVisibilityType.Visible;

                // Prepare output stream
                var outputStream = new MemoryStream();

                // Save the document
                targetWorkbook.SaveDocument(outputStream, outputFormat);

                outputStream.Position = 0;
                documentStreams.Add(outputStream);
            }
            return documentStreams.AsReadOnly();
        }
        #endregion Code
    }
}
