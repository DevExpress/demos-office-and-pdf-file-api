#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentProtection {
    [ApiController]
    [Route("api/presentation")]
    public class PresentationEncryptPasswordController : PresentationControllerBase {
        [HttpPost("encrypt-password/encrypt")]
        public async Task<IActionResult> EncryptDocument(IFormFile file, [FromForm] PresentationEncryptPasswordModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            EncryptionType encryptionTypeConverted = GetEncryptionType(model.EncryptionType);
            DocumentFormat outputFormatConverted = PresentationFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentStream = EncryptDocumentWithPassword(inputStream, outputFormatConverted, model.Password, encryptionTypeConverted);
            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }

        EncryptionType GetEncryptionType(string encryptionType) {
            return encryptionType.ToLowerInvariant() switch {
                "compatible" => EncryptionType.Compatible,
                "strong" => EncryptionType.Strong,
                _ => EncryptionType.Strong
            };
        }

        #region Code
        Stream EncryptDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, EncryptionType encryptionType) {
            using var presentation = new Presentation(inputStream);

            var options = new EncryptionOptions(password, encryptionType);
            presentation.Encrypt(options);

            var outputStream = new MemoryStream();
            presentation.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
