#region Code

using DevExpress.Docs;
using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentProtection {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordSecureDocumentLoadingController : WordControllerBase {
        [HttpPost("secure-document-loading/load")]
        public async Task<IActionResult> Load(IFormFile file, [FromForm] WordSecureDocumentLoadingModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            // Process document and remove unsafe content
            var outputStream = TryLoadDocument(
                inputStream,
                new DocumentSecurityLoadingLimits() {
                    MaxSectionsCount = model.MaxSectionsCount,
                    MaxParagraphCount = model.MaxParagraphCount,
                    MaxTableCount = model.MaxTableCount,
                    MaxTableRowCount = model.MaxTableRowCount,
                    MaxSubDocumentCount = model.MaxSubDocumentCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxXmlElementCount = 5000,
                    MaxXmlElementDepth = 50
                },
                new DocumentSecurityLoadingOptions() {
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveCustomXMLParts = model.RemoveCustomXMLParts,
                    RemoveMacros = model.RemoveMacros,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemoveExternalImages = model.RemoveExternalImages,
                    RemoveIncludePictureFields = model.RemoveIncludePictureFields,
                    RemoveDDEFields = model.RemoveDDEFields,
                    RestrictedHyperlinkRemovalMode = model.RestrictedHyperlinkRemovalMode
                },
                out var loadErrorMessage,
                outputFormat);

            if(string.IsNullOrEmpty(loadErrorMessage)) {
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
            else {
                return BadRequest(loadErrorMessage);
            }
        }

        #region Code
        Stream TryLoadDocument(Stream inputStream, DocumentSecurityLoadingLimits securityLoadingLimits,
                               DocumentSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
            var loadOptions = new LoadOptions() {
                SecurityLoadingHandler = new DocumentSecurityLoadingHandler(),
                SecurityLoadingOptions = securityLoadingOptions,
                SecurityLoadingLimits = securityLoadingLimits
            };

            loadErrorMessage = "";
            try {
                using var document = new WordDocument(inputStream, loadOptions);

                // Save document
                var outputStream = new MemoryStream();

                document.Save(outputStream, outputFormat);

                outputStream.Position = 0;
                return outputStream;
            }
            catch(SecurityLoadingLimitExceededException limitExceededException) {
                loadErrorMessage = limitExceededException.Message;
                return new MemoryStream();
            }
        }

        class DocumentSecurityLoadingHandler : ISecurityLoadingHandler {
            public SecurityLimitAction HandleLimitExceeded(SecurityLimitExceededParams parameters) {
                // Use the HandleLimitExceeded method to log information about the security loading limits
                // Return SecurityLimitAction.Continue to load a document even if it exceeds the specified limits
                return SecurityLimitAction.Throw;
            }

            public SecurityViolationAction HandleOptionsViolation(SecurityOptionsViolationParams parameters) {
                // Use the HandleOptionsViolation method to log information about the security loading options violation
                // Return SecurityViolationAction.Keep to ignore the violation and load a document even if it contains unsafe content
                return SecurityViolationAction.Remove;
            }
        }
        #endregion Code
    }
}