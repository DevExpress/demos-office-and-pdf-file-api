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
    public class ExcelProtectWorksheetController : SpreadsheetControllerBase {
        [HttpPost("protect-password/protect-worksheet")]
        public async Task<IActionResult> Protect(IFormFile file, [FromForm] ExcelProtectWorksheetModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            DocumentFormat outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var worksheetRange = RangeHelper.ParseRange(model.WorksheetRange);

            Stream outputStream = ProtectWorksheetWithPassword(
                inputStream,
                outputFormat,
                model.WorksheetPassword,
                worksheetRange.ToArray(),
                model.SelectedWorksheetProtectionPermissions
                );
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ProtectWorksheetWithPassword(Stream inputStream, DocumentFormat outputFormat, string worksheetPassword,
            int[] worksheetIndexes, IEnumerable<WorksheetProtectionPermissions> worksheetProtectionPermissions) {
            using var workbook = new Workbook(inputStream);

            var protectionOptions = new WorksheetProtectionOptions(worksheetPassword) {
                Permissions = WorksheetProtectionPermissions.None
            };
            foreach(WorksheetProtectionPermissions value in worksheetProtectionPermissions)
                protectionOptions.Permissions |= value;

            int worksheetCount = workbook.Worksheets.Count;
            foreach(int index in worksheetIndexes) {
                if(index < worksheetCount)
                    workbook.Worksheets[index].Protect(protectionOptions);
            }

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
