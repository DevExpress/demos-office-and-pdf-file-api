#region Code

using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using System.Text;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentProtection {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordInspectAndSanitizeController : WordControllerBase {
        [HttpPost("secure-document-loading/inspect")]
        public async Task<IActionResult> Inspect(IFormFile file, [FromForm] WordInspectAndSanitizeContentModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Inspect Document
            var inspectionResult = InspectDocument(inputStream, DocumentInspectOptions.All);
            using var outputStream = new MemoryStream();
            using(var writer = new StreamWriter(outputStream, Encoding.UTF8)) {
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

            var outputFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Sanitize Document
            using var outputStream = SanitizeDocument(
                inputStream,
                outputFormat,
                new DocumentSanitizeOptions() {
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
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        string InspectDocument(Stream inputStream, DocumentInspectOptions inspectOptions) {
            using var document = new WordDocument(inputStream);

            var output = new StringBuilder();

            var inspectResult = document.Inspect(inspectOptions);
            if(inspectResult.ContentTypes.Count == 0) {
                output.AppendLine("No issues found.");
            }
            else {
                output.AppendLine("Document contains the following hidden/sensitive content:");
                output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
            }

            return output.ToString();
        }

        Stream SanitizeDocument(Stream inputStream, DocumentFormat outputFormat, DocumentSanitizeOptions sanitizeOptions = null) {
            using var document = new WordDocument(inputStream);

            // Inspect a document and create DocumentSanitizeOptions based on the inspection result
            if(sanitizeOptions == null) {
                var inspectResult = document.Inspect();
                sanitizeOptions = inspectResult.CreateSanitizeOptions();
            }

            // Use the sanitizeResult object to get information about the removed content
            var sanitizeResult = document.Sanitize(sanitizeOptions);

            // Save document
            var outputStream = new MemoryStream();

            document.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}