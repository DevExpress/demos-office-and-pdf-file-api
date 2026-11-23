#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
using DevExpress.Docs;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentProtection {
    [ApiController]
    [Route("api/presentation")]
    public class PresentationSecureDocumentLoadingController : PresentationControllerBase {
        [HttpPost("secure-document-loading/load")]
        public async Task<IActionResult> Load(IFormFile file, [FromForm] PresentationSecureDocumentLoadingModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputFormat = PresentationFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process presentation and remove unsafe content
            string loadErrorMessage = string.Empty;

            var outputStream = TryLoadPresentation(
                inputStream,
                new PresentationSecurityLoadingLimits() {
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxShapeCountPerSlide = model.MaxShapeCountPerSlide,
                    MaxSlideCount = model.MaxSlideCount,
                    MaxXmlElementCount = 1000,
                    MaxXmlElementDepth = 20
                },
                new PresentationSecurityLoadingOptions() {
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveCustomXMLParts = model.RemoveCustomXMLParts,
                    RemoveEmbeddedChartData = model.RemoveEmbeddedChartData,
                    RemoveExternalImages = model.RemoveExternalImages,
                    RemoveMacros = model.RemoveMacros,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RestrictedHyperlinkRemovalMode = model.RestrictedHyperlinkRemovalMode
                },
                out loadErrorMessage,
                outputFormat);

            if(string.IsNullOrEmpty(loadErrorMessage)) {
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
            else {
                return BadRequest(loadErrorMessage);
            }
        }
        #region Code
        Stream TryLoadPresentation(Stream inputStream, PresentationSecurityLoadingLimits securityLoadingLimits,
            PresentationSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
            var loadOptions = new LoadOptions() {
                // Specify ISecurityLoadingHandler to handle security loading limits and options violations
                SecurityLoadingHandler = new PresentationSecurityLoadingHandler(),
                SecurityLoadingOptions = securityLoadingOptions,
                SecurityLoadingLimits = securityLoadingLimits
            };

            loadErrorMessage = "";
            try {
                using var presentation = new Presentation(inputStream, loadOptions);

                // Save presentation
                var outputStream = new MemoryStream();

                presentation.Save(outputStream, outputFormat);

                outputStream.Position = 0;
                return outputStream;
            }
            catch(SecurityLoadingLimitExceededException limitExceededException) {
                loadErrorMessage = limitExceededException.Message;
                return new MemoryStream();
            }
        }

        // Implement ISecurityLoadingHandler to handle security loading limits and options violations
        class PresentationSecurityLoadingHandler : ISecurityLoadingHandler {
            public SecurityLimitAction HandleLimitExceeded(SecurityLimitExceededParams parameters) {
                // Use the HandleLimitExceeded method to log information about the security loading limits.
                // Return the "SecurityLimitAction.Continue" value to ignore the violation and load a presentation even if it exceeds the specified limits
                return SecurityLimitAction.Throw;
            }

            public SecurityViolationAction HandleOptionsViolation(SecurityOptionsViolationParams parameters) {
                // Use the HandleOptionsViolation method to log information about the security loading options violations.
                // Return the "SecurityViolationAction.Keep" value to ignore the violation and load a presentation even if it contains unsafe content
                return SecurityViolationAction.Remove;
            }
        }
        #endregion Code
    }
}
