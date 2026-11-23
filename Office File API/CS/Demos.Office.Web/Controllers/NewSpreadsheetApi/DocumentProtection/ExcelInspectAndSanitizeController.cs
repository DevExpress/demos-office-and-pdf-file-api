#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Docs.Office;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentProtection {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelInspectAndSanitizeController : SpreadsheetControllerBase {
        [HttpPost("secure-document-loading/inspect")]
        public async Task<IActionResult> Inspect(IFormFile file, [FromForm] ExcelInspectAndSanitizeContentModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            string inspectionResult = InspectWorkbook(inputStream, WorkbookInspectOptions.All);
            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(inspectionResult);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
            }
        }
        [HttpPost("secure-document-loading/sanitize")]
        public async Task<IActionResult> Sanitize(IFormFile file, [FromForm] ExcelInspectAndSanitizeContentModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            using var outputStream = SanitizeWorkbook(
                inputStream,
                outputFormat,
                new WorkbookSanitizeOptions() {
                    HiddenColumns = model.RemoveHiddenColumns ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenRows = model.RemoveHiddenRows ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenSheets = model.RemoveHiddenSheets ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    InvisibleCellText = model.RemoveInvisibleCellText ? InvisibleContentSanitizeMode.Remove : InvisibleContentSanitizeMode.Ignore,
                    RemoveMacros = model.RemoveMacros,
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemoveThreadedComments = model.RemoveThreadedComments,
                    RemoveComments = model.RemoveComments,
                    RemoveCustomXmlParts = model.RemoveCustomXMLParts,
                    Metadata = model.RemoveMetadata ? MetadataRemovalScope.All : MetadataRemovalScope.None,
                });
            return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
        }
        #region Code
        string InspectWorkbook(Stream inputStream, WorkbookInspectOptions inspectOptions) {
            using var workbook = new Workbook(inputStream);

            StringBuilder output = new StringBuilder();

            var inspectResult = workbook.Inspect(inspectOptions);
            if(inspectResult.ContentTypes.Count == 0) {
                output.AppendLine("No issues found.");
            }
            else {
                output.AppendLine("Document contains the following hidden/sensitive content:");
                output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
            }
            return output.ToString();
        }

        Stream SanitizeWorkbook(Stream inputStream, DocumentFormat outputFormat, WorkbookSanitizeOptions sanitizeOptions = null) {
            using var workbook = new Workbook(inputStream);

            if(sanitizeOptions == null) {
                var inspectResult = workbook.Inspect();
                sanitizeOptions = inspectResult.CreateSanitizeOptions();
            }

            var sanitizeResult = workbook.Sanitize(sanitizeOptions);

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
