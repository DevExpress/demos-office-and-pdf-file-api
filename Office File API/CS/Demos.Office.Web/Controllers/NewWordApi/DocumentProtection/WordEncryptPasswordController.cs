#region Code

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
    public class WordEncryptPasswordController : WordControllerBase {
        [HttpPost("encrypt-password/encrypt")]
        public async Task<IActionResult> EncryptDocument(IFormFile file, [FromForm] WordEncryptPasswordModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var documentStream = EncryptDocumentWithPassword(inputStream, outputFormat, model.Password, model.EncryptionType);

            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream EncryptDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, EncryptionType encryptionType) {
            // Load document from stream
            using var document = new WordDocument(inputStream);

            // Create encryption settings
            document.Encrypt(new EncryptionOptions(password, encryptionType));

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the encrypted document
            document.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}