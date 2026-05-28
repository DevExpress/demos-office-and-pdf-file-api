#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentProtection;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentProtection {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordProtectPasswordController : WordControllerBase {
        [HttpPost("protect-password/protect")]
        public async Task<IActionResult> ProtectDocument(IFormFile file, [FromForm] WordProtectPasswordModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormatConverted = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = ProtectDocumentWithPassword(inputStream, outputFormatConverted, model.Password, model.ProtectionType);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProtectDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, DocumentProtectionType protectionType) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(inputStream);

            // Get document reference
            var document = wordProcessor.Document;

            // Unprotect document if it's already protected
            if(document.IsDocumentProtected)
                document.Unprotect();

            if(protectionType == DocumentProtectionType.FillInForms)
                foreach(Section section in document.Sections)
                    section.ProtectedForForms = true;

            // Apply document protection with password
            document.Protect(password, protectionType);

            // Save protected document
            var outputStream = new MemoryStream();

            wordProcessor.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
