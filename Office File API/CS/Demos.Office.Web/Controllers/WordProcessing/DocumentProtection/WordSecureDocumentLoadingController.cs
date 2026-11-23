#region Code
using DevExpress.XtraRichEdit;
using DevExpress.Office;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentProtection {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordSecureDocumentLoadingController : WordControllerBase {
        [HttpPost("secure-document-loading/load")]
        public async Task<IActionResult> Load(IFormFile file, [FromForm] WordSecureDocumentLoadingModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            // Process document and remove unsafe content
            string loadErrorMessage = string.Empty;
            var outputStream = TryLoadDocument(
                inputStream,
                new WordProcessingSecurityLoadingLimits() {
                    MaxSectionsCount = model.MaxSectionsCount,
                    MaxParagraphCount = model.MaxParagraphCount,
                    MaxTableCount = model.MaxTableCount,
                    MaxTableRowCount = model.MaxTableRowCount,
                    MaxSubDocumentCount = model.MaxSubDocumentCount,
                    MaxFileSize = 5 * 1024 * 1024,
                    MaxXmlElementCount = 5000,
                    MaxXmlElementDepth = 50
                },
                new WordProcessingSecurityLoadingOptions() {
                    RemoveActiveXContent = model.RemoveActiveXContent,
                    RemoveCustomXMLParts = model.RemoveCustomXMLParts,
                    RemoveMacros = model.RemoveMacros,
                    RemoveOleObjects = model.RemoveOleObjects,
                    RemoveExternalImages = model.RemoveExternalImages,
                    RemoveIncludePictureFields = model.RemoveIncludePictureFields,
                    RemoveDDEFields = model.RemoveDDEFields,
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
        Stream TryLoadDocument(Stream inputStream, WordProcessingSecurityLoadingLimits securityLoadingLimits,
            WordProcessingSecurityLoadingOptions securityLoadingOptions, out string loadErrorMessage, DocumentFormat outputFormat) {
            using var richEditDocumentServer = new RichEditDocumentServer();

            var richSecurityLoadingOptions = richEditDocumentServer.Options.SecurityLoadingOptions;
            richSecurityLoadingOptions.RemoveActiveXContent = securityLoadingOptions.RemoveActiveXContent;
            richSecurityLoadingOptions.RemoveCustomXMLParts = securityLoadingOptions.RemoveCustomXMLParts;
            richSecurityLoadingOptions.RemoveMacros = securityLoadingOptions.RemoveMacros;
            richSecurityLoadingOptions.RemoveOleObjects = securityLoadingOptions.RemoveOleObjects;
            richSecurityLoadingOptions.RestrictedHyperlinkRemovalMode = securityLoadingOptions.RestrictedHyperlinkRemovalMode;
            richSecurityLoadingOptions.RemoveExternalImages = securityLoadingOptions.RemoveExternalImages;
            richSecurityLoadingOptions.RemoveIncludePictureFields = securityLoadingOptions.RemoveIncludePictureFields;
            richSecurityLoadingOptions.RemoveDDEFields = securityLoadingOptions.RemoveDDEFields;

            var richSecurityLoadingLimits = richEditDocumentServer.Options.SecurityLoadingLimits;
            richSecurityLoadingLimits.MaxFileSize = securityLoadingLimits.MaxFileSize;
            richSecurityLoadingLimits.MaxXmlElementCount = securityLoadingLimits.MaxXmlElementCount;
            richSecurityLoadingLimits.MaxXmlElementDepth = securityLoadingLimits.MaxXmlElementDepth;
            richSecurityLoadingLimits.MaxSectionsCount = securityLoadingLimits.MaxSectionsCount;
            richSecurityLoadingLimits.MaxParagraphCount = securityLoadingLimits.MaxParagraphCount;
            richSecurityLoadingLimits.MaxTableCount = securityLoadingLimits.MaxTableCount;
            richSecurityLoadingLimits.MaxTableRowCount = securityLoadingLimits.MaxTableRowCount;
            richSecurityLoadingLimits.MaxSubDocumentCount = securityLoadingLimits.MaxSubDocumentCount;

            richEditDocumentServer.SecurityLoadingLimitExceeded += (o, ea) => {
                // Use the SecurityLoadingLimitExceeded event to log information about the security loading limits
                // Set the "Handled" parameter to true to load a document even if it exceeds the specified limits
            };

            richEditDocumentServer.SecurityLoadingOptionsViolation += (o, ea) => {
                // Use the SecurityLoadingOptionsViolation event to log information about the security loading options violation
                // Set the "Handled" parameter to true to ignore the violation and load a document even if it contains unsafe content
            };

            loadErrorMessage = "";
            try {
                richEditDocumentServer.LoadDocument(inputStream);

                // Save document
                var outputStream = new MemoryStream();

                richEditDocumentServer.SaveDocument(outputStream, outputFormat);

                outputStream.Position = 0;
                return outputStream;
            }
            catch(SecurityLoadingLimitExceededException limitExceededException) {
                loadErrorMessage = limitExceededException.Message;
                return new MemoryStream();
            }
        }
        #endregion Code
    }
}

