#region Code
using DevExpress.XtraRichEdit;
using DevExpress.Office;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentProtection {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordInspectAndSanitizeController : WordControllerBase {
        [HttpPost("secure-document-loading/inspect")]
        public async Task<IActionResult> Inspect(IFormFile file, [FromForm] WordInspectAndSanitizeContentModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Inspect Document
            string inspectionResult = InspectDocument(inputStream, WordProcessingInspectOptions.All);
            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(inspectionResult);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
            }
        }
        [HttpPost("secure-document-loading/sanitize")]
        public async Task<IActionResult> Sanitize(IFormFile file, [FromForm] WordInspectAndSanitizeContentModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Sanitize Document
            using var outputStream = SanitizeDocument(
                inputStream,
                outputFormat,
                new WordProcessingSanitizeOptions() {
                    HiddenText = model.RemoveHiddenText ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    InvisibleText = model.RemoveInvisibleText ? InvisibleContentSanitizeMode.Remove : InvisibleContentSanitizeMode.Ignore,
                    TrackedChanges = model.RemoveTrackChanges ? TrackedChangesSanitizeMode.Reject : TrackedChangesSanitizeMode.Ignore,
                    RemoveMacros = model.RemoveMacros,
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemoveComments = model.RemoveComments,
                    RemoveCustomXmlParts = model.RemoveCustomXMLParts,
                    Metadata = model.RemoveMetadata ? MetadataRemovalScope.All : MetadataRemovalScope.None,
                });
            return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
        }
        #region Code
        string InspectDocument(Stream inputStream, WordProcessingInspectOptions inspectOptions) {
            using var richEditDocumentServer = new RichEditDocumentServer();
            richEditDocumentServer.LoadDocument(inputStream);

            StringBuilder output = new StringBuilder();

            var inspectResult = richEditDocumentServer.Inspect(inspectOptions);
            if(inspectResult.ContentTypes.Count == 0) {
                output.AppendLine("No issues found.");
            }
            else {
                output.AppendLine("Document contains the following hidden/sensitive content:");
                output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
            }
            return output.ToString();
        }

        Stream SanitizeDocument(Stream inputStream, DocumentFormat outputFormat, WordProcessingSanitizeOptions sanitizeOptions = null) {
            using var richEditDocumentServer = new RichEditDocumentServer();
            richEditDocumentServer.LoadDocument(inputStream);

            // Inspect a document and create WordProcessingSanitizeOptions based on the inspection result
            if(sanitizeOptions == null) {
                var inspectResult = richEditDocumentServer.Inspect();
                sanitizeOptions = inspectResult.CreateSanitizeOptions();
            }

            // Use the sanitizeResult object to get information about the removed content
            var sanitizeResult = richEditDocumentServer.Sanitize(sanitizeOptions);

            // Save document
            var outputStream = new MemoryStream();

            richEditDocumentServer.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
