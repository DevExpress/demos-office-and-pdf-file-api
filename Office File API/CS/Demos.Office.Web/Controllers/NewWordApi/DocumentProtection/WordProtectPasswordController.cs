#region Code

using DevExpress.Docs.Word;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentProtection {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordProtectPasswordController : WordControllerBase {
        [HttpPost("protect-password/protect")]
        public async Task<IActionResult> ProtectDocument(IFormFile file, [FromForm] WordProtectPasswordModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var outputStream = ProtectDocumentWithPassword(inputStream, outputFormat, model.Password, model.ProtectionType);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProtectDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, DocumentProtectionType protectionType) {
            using var document = new WordDocument(inputStream);

            // Unprotect document if it's already protected
            if(document.ProtectionType != DocumentProtectionType.None)
                document.RemoveProtection();

            // Apply document protection with password
            var protectionOptions = new DocumentProtectionOptions(password) {
                ProtectionType = protectionType
            };
            document.Protect(protectionOptions);

            // Save protected document
            var outputStream = new MemoryStream();

            document.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}