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
    public class ExcelProtectWorksheetController : SpreadsheetControllerBase {
        [HttpPost("protect-password/protect-worksheet")]
        public async Task<IActionResult> Protect(IFormFile file, [FromForm] ExcelProtectWorksheetModel model) {
            // Load document from stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Apply document protection
            Stream outputStream = ProtectWorksheetWithPassword(
                inputStream,
                outputFormat,
                model.WorksheetPassword,
                model.SheetProtectionMode == SheetProtectionMode.ActiveSheet,
                model.SelectedWorksheetProtectionPermissions
                );
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProtectWorksheetWithPassword(Stream inputStream, DocumentFormat outputFormat, string worksheetPassword,
            bool protectActiveWorksheetOnly, IEnumerable<WorksheetProtectionPermissions> worksheetProtectionPermissions) {
            using var workbook = new Workbook();

            workbook.LoadDocument(inputStream);

            WorksheetProtectionPermissions permissions = 0;
            foreach(WorksheetProtectionPermissions value in worksheetProtectionPermissions)
                permissions |= value;

            // Get worksheets to protect
            var worksheetsToProtect = new List<Worksheet>();
            if(protectActiveWorksheetOnly) {
                var activeSheet = workbook.Worksheets.ActiveWorksheet;
                worksheetsToProtect.Add(activeSheet);
            } else
                worksheetsToProtect = workbook.Worksheets.ToList();

            // Protect worksheets
            foreach(var worksheet in worksheetsToProtect) {
                worksheet.Protect(worksheetPassword ?? string.Empty, permissions);
            }

            // Save protected document
            var outputStream = new MemoryStream();

            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
