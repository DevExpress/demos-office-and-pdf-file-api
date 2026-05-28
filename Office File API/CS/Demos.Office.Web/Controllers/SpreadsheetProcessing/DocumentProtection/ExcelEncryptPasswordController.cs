#region Code
using DevExpress.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DocumentProtection {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelEncryptPasswordController : SpreadsheetControllerBase {
        [HttpPost("encrypt-password/encrypt")]
        public async Task<IActionResult> EncryptDocument(IFormFile file, [FromForm] ExcelEncryptPasswordModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            EncryptionType encryptionTypeConverted = GetEncryptionType(model.EncryptionType);
            DocumentFormat outputFormatConverted = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Encrypt the document with a password
            var documentStream = EncryptDocumentWithPassword(inputStream, outputFormatConverted, model.Password, encryptionTypeConverted);
            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }

        EncryptionType GetEncryptionType(string encryptionType) {
            return encryptionType.ToLowerInvariant() switch {
                "compatible" => EncryptionType.Compatible,
                "strong" => EncryptionType.Strong,
                _ => EncryptionType.Strong // Default to Strong if unknown
            };
        }

        #region Code
        Stream EncryptDocumentWithPassword(Stream inputStream, DocumentFormat outputFormat, string password, EncryptionType encryptionType) {
            using var workboook = new Workbook();

            // Load document from stream
            workboook.LoadDocument(inputStream);

            // Create encryption settings
            var encryptionSettings = new EncryptionSettings {
                Password = password,
                Type = encryptionType
            };

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the encrypted document
            workboook.SaveDocument(outputStream, outputFormat, encryptionSettings);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
