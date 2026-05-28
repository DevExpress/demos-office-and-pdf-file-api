using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentProtection {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfEncryptPasswordController : PdfControllerBase {
        [HttpPost("document-protection/encrypt-password")]
        public async Task<IActionResult> EncryptPdfWithPassword(
            IFormFile file,
            [FromForm] PdfEncryptPasswordModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = Encrypt(
                documentStream,
                model.EncryptionAlgorithm,
                model.DocumentDataExtractionPermissions,
                model.DocumentInteractivityPermissions,
                model.DocumentModificationPermissions,
                model.DocumentPrintingPermissions,
                model.OwnerPassword,
                model.UserPassword);
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream Encrypt(Stream documentStream, EncryptionAlgorithm algorithm,
            DocumentDataExtractionPermissions dataExtractionPermissions, DocumentInteractivityPermissions interactivityPermissions,
            DocumentModificationPermissions modificationPermissions, DocumentPrintPermissions printPermissions,
            string ownerPassword, string userPassword) {
            using var doc = new PdfDocument(documentStream);

            doc.Encrypt(new EncryptionOptions(ownerPassword, userPassword) {
                Algorithm = algorithm,
                DataExtractionPermissions = dataExtractionPermissions,
                InteractivityPermissions = interactivityPermissions,
                ModificationPermissions = modificationPermissions,
                PrintPermissions = printPermissions
            });

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
