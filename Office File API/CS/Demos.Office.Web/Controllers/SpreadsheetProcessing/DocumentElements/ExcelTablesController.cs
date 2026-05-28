#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelTablesController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("tables/generate")]
        public async Task<IActionResult> GenerateTablesWorkbook([FromForm] ExcelTablesModel model) {
            // Determine output format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with table
            var outputStream = GenerateDocumentWithTable(model.NumberOfRows, model.NumberOfColumns, true, model.CreateCustomTableStyle, model.ShowTotals, outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream GenerateDocumentWithTable(int rowCount, int columnCount, bool hasHeaders, bool useCustomTableStyle, bool showTotals, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            // Get first worksheet in the workbook
            var worksheet = workbook.Worksheets.First();

            // Specify a cell range that contains table data 
            var range = worksheet.Range.FromLTRB(0, 0, columnCount - 1, hasHeaders ? rowCount : rowCount - 1);

            // Populate the header row with custom text
            if(hasHeaders) {
                for(int column = 0; column < columnCount; column++) {
                    var cell = range[0, column];
                    var cellText = $"Header {column + 1}";
                    cell.SetValue(cellText);
                }
            }

            // Populate the table range with sample data
            int startRow = hasHeaders ? 1 : 0;
            for(int row = 0; row < rowCount; row++) {
                for(int column = 0; column < columnCount; column++) {
                    var cell = range[startRow + row, column];
                    var cellText = column == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{column + 1}";
                    cell.SetValue(cellText);
                }
            }

            // Insert a table
            var table = worksheet.Tables.Add(range, hasHeaders);

            // Access the workbook's collection of table styles
            var tableStyles = workbook.TableStyles;

            // Apply table style
            if(useCustomTableStyle) {
                // Create a custom table style
                table.Style = CreateCustomTableStyle(tableStyles, "CustomTableStyle");
            }
            else {
                // Access the built-in table style by its Id
                table.Style = tableStyles[BuiltInTableStyleId.TableStyleDark9];
            }

            if(showTotals) {
                table.ShowTotals = true;
                table.Columns.Last().TotalRowFunction = TotalRowFunction.Count;
            }

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        TableStyle CreateCustomTableStyle(TableStyleCollection tableStyles, string styleName) {
            // Add a new style under the specified name to the table style collection
            var customTableStyle = tableStyles.Add(styleName);

            // Modify table style formatting
            // Specify format characteristics for different table elements
            customTableStyle.BeginUpdate();
            try {
                customTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(107, 107, 107);

                // Format the header row
                var headerRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.HeaderRow];
                headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166);
                headerRowStyle.Font.Color = Color.White;
                headerRowStyle.Font.Bold = true;

                // Format the total row
                var totalRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.TotalRow];
                totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211);
                totalRowStyle.Font.Color = Color.White;
                totalRowStyle.Font.Bold = true;

                // Specify banded row formatting for the table
                var secondRowStripeStyle = customTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe];
                secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234);
                secondRowStripeStyle.StripeSize = 1;
            }
            finally {
                customTableStyle.EndUpdate();
            }
            return customTableStyle;
        }
        #endregion Code
    }
}
