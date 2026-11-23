#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Docs.Office;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentProtection {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelEncryptPasswordController : SpreadsheetControllerBase {
        [HttpPost("encrypt-password/encrypt")]
        public async Task<IActionResult> EncryptDocument(IFormFile file, [FromForm] ExcelEncryptPasswordModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            EncryptionType encryptionTypeConverted = GetEncryptionType(model.EncryptionType);
            DocumentFormat outputFormatConverted = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

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
            using var workbook = new Workbook(inputStream);

            var encryptionOptions = new EncryptionOptions(password, encryptionType);
            workbook.Encrypt(encryptionOptions);

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
