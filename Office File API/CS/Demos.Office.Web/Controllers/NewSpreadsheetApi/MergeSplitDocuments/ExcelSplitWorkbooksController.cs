#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.MergeSplitDocuments;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.MergeSplitDocuments {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelSplitWorkbooksController : SpreadsheetControllerBase {
        [HttpPost("split-workbooks/split")]
        public async Task<IActionResult> Split(IFormFile file, [FromForm] ExcelSplitWorkbooksModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;
            DocumentFormat outputFormatConverted = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentStreams = SplitWorkbook(model.SplitMode == VisibilityMode.VisibleOnly, inputStream, outputFormatConverted);

            if(documentStreams == null || documentStreams.Count == 0)
                return StatusCode(500, new { Success = false, Message = "Failed to split workbook" });

            if(documentStreams.Count == 1) {
                var fileName = "document";
                return CreateOutputFile((Stream)documentStreams[0], model.OutputFormat, fileName);
            } else {
                var zipFileName = $"{model.OutputName}.zip";
                return CreateZipOutputFile(documentStreams, model.OutputFormat, "worksheet", zipFileName);
            }
        }
        #region Code
        IReadOnlyList<Stream> SplitWorkbook(bool splitVisibleOnly, Stream inputStream, DocumentFormat outputFormat) {
            using var sourceWorkbook = new Workbook(inputStream);

            var options = new WorkbookExtractionOptions { UnhideSheets = true };

            var documentStreams = new List<Stream>();
            for(var index = 0; index < sourceWorkbook.Sheets.Count; index++) {
                if(!ShouldExtractSheet(splitVisibleOnly, sourceWorkbook.Sheets[index]))
                    continue;

                using var targetWorkbook = sourceWorkbook.Extract(options, index);

                targetWorkbook.Calculate();

                var outputStream = new MemoryStream();

                targetWorkbook.Save(outputStream, outputFormat);

                outputStream.Position = 0;
                documentStreams.Add(outputStream);
            }
            return documentStreams.AsReadOnly();
        }

        static bool ShouldExtractSheet(bool splitVisibleOnly, Sheet sheet) {
            if(sheet.Type == SheetType.ChartSheet)
                return false;
            if(splitVisibleOnly)
                return sheet.Visibility == SheetVisibility.Visible;
            return sheet.Visibility != SheetVisibility.VeryHidden;
        }
        #endregion Code
    }
}
