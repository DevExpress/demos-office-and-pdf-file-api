#region Code
using DevExpress.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelGroupDataController : SpreadsheetControllerBase {
        [HttpPost("group-data/group")]
        public async Task<IActionResult> Group(IFormFile file, [FromForm] ExcelGroupDataModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Group data on the first worksheet
            Stream outputStream = GroupData(
                        inputStream,
                        model.GroupMode == GroupMode.Rows || model.GroupMode == GroupMode.Both,
                        model.GroupMode == GroupMode.Columns || model.GroupMode == GroupMode.Both,
                        model.ShowSubtotals,
                        model.CollapseGroups,
                        targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream GroupData(Stream inputStream, bool groupRows, bool groupColumns, bool showSubtotals, bool collapsed, DocumentFormat outputFormat) {

            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Obtain the first worksheet
            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Remove existing subtotals before creating groups
            if(!showSubtotals) {
                firstWorksheet.RemoveSubtotal(firstWorksheet["B3:G35"]);
            }

            // Group rows and columns  
            if(groupRows) {
                // Create Quarterly total groups
                firstWorksheet.Rows.Group(3, 8, collapsed);
                firstWorksheet.Rows.Group(showSubtotals ? 11 : 10, showSubtotals ? 16 : 15, collapsed);
                firstWorksheet.Rows.Group(showSubtotals ? 19 : 17, showSubtotals ? 24 : 22, collapsed);
                firstWorksheet.Rows.Group(showSubtotals ? 27 : 24, showSubtotals ? 32 : 29, collapsed);

                // Create Grand Total group only if subtotals are shown
                if(showSubtotals)
                    firstWorksheet.Rows.Group(2, 33, collapsed);
            }

            if(groupColumns) {
                // Group columns  
                firstWorksheet.Columns.Group("C", "F", collapsed);
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
