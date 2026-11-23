#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentProtection {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelProtectWorkbookController : SpreadsheetControllerBase {
        [HttpPost("protect-password/protect-workbook")]
        public async Task<IActionResult> ProtectDocument(IFormFile file, [FromForm] ExcelProtectWorkbookModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormatConverted = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = ProtectWorkbookWithPassword(
                inputStream,
                outputFormatConverted,
                model.WorkbookPassword,
                model.LockStructure,
                model.LockWindows
                );
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProtectWorkbookWithPassword(Stream inputStream, DocumentFormat outputFormat, string workbookPassword, bool lockStructure, bool lockWindows) {
            using var workbook = new Workbook(inputStream);

            if(workbook.ProtectionMode == WorkbookProtectionMode.None) {
                workbook.Protect(new WorkbookProtectionOptions(workbookPassword) {
                    LockStructure = lockStructure,
                    LockWindows = lockWindows,
                    RequirePasswordToModify = true
                });
            }

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
