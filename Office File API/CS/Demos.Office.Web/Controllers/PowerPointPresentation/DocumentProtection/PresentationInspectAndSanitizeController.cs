#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentProtection {
    [ApiController]
    [Route("api/presentation")]
    public class PresentationInspectAndSanitizeController : PresentationControllerBase {
        [HttpPost("secure-document-loading/inspect")]
        public async Task<IActionResult> Inspect(IFormFile file, [FromForm] PresentationInspectAndSanitizeContentModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Inspect Presentation
            string inspectionResult = InspectPresentation(inputStream, PresentationInspectOptions.All);
            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(inspectionResult);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
            }
        }
        [HttpPost("secure-document-loading/sanitize")]
        public async Task<IActionResult> Sanitize(IFormFile file, [FromForm] PresentationInspectAndSanitizeContentModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputFormat = PresentationFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Sanitize Presentation
            using var outputStream = SanitizePresentation(
                inputStream,
                outputFormat,
                new PresentationSanitizeOptions() {
                    HiddenSlides = model.RemoveHiddenSlides ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,
                    HiddenShapes = model.RemoveHiddenShapes ? HiddenContentSanitizeMode.Remove : HiddenContentSanitizeMode.Ignore,

                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveComments = model.RemoveComments,
                    RemoveCustomXmlParts = model.RemoveCustomXMLParts,
                    RemoveEmbeddedChartData = model.RemoveEmbeddedChartData,
                    RemoveMacros = model.RemoveMacros,
                    Metadata = model.RemoveMetadata ? MetadataRemovalScope.All : MetadataRemovalScope.None,
                    RemoveNotes = model.RemoveNotes,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemoveOffSlideContent = model.RemoveOffSlideContent,
                });
            return CreateOutputFile(outputStream, model.OutputFileFormat, model.OutputFileName);
        }
        #region Code
        string InspectPresentation(Stream inputStream, PresentationInspectOptions inspectOptions) {
            using var presentation = new Presentation(inputStream);

            StringBuilder output = new StringBuilder();

            var inspectResult = presentation.Inspect(inspectOptions);
            if(inspectResult.ContentTypes.Count == 0) {
                output.AppendLine("No issues found.");
            }
            else {
                output.AppendLine("Presentation contains the following hidden/sensitive content:");
                output.AppendLine(string.Join(Environment.NewLine, inspectResult.ContentTypes));
            }
            return output.ToString();
        }

        Stream SanitizePresentation(Stream inputStream, DocumentFormat outputFormat, PresentationSanitizeOptions sanitizeOptions = null) {
            using var presentation = new Presentation(inputStream);

            // Inspect a presentation and create PresentationSanitizeOptions based on the inspection result
            if(sanitizeOptions == null) {
                var inspectResult = presentation.Inspect();
                sanitizeOptions = inspectResult.CreateSanitizeOptions();
            }

            // Use the sanitizeResult object to get information about the removed content
            var sanitizeResult = presentation.Sanitize(sanitizeOptions);

            // Save presentation
            var outputStream = new MemoryStream();

            presentation.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
