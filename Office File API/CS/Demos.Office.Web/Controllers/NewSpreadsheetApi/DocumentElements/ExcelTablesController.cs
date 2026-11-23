#region Code
using DevExpress.Docs.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentElements {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelTablesController : SpreadsheetControllerBase {
        [HttpPost("tables/generate")]
        public async Task<IActionResult> GenerateTablesWorkbook([FromForm] ExcelTablesModel model) {
            var outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var outputStream = GenerateDocumentWithTable(model.NumberOfRows, model.NumberOfColumns, true, model.CreateCustomTableStyle, model.ShowTotals, outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream GenerateDocumentWithTable(int rowCount, int columnCount, bool hasHeaders, bool useCustomTableStyle, bool showTotals, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            var worksheet = workbook.Worksheets.First();

            int startRow = hasHeaders ? 1 : 0;
            var range = worksheet.GetRange(0, 0, startRow + rowCount, columnCount);

            if(hasHeaders) {
                for(int column = 0; column < columnCount; column++)
                    worksheet.Cells[0, column].Value = $"Header {column + 1}";
            }

            for(int row = 0; row < rowCount; row++) {
                for(int column = 0; column < columnCount; column++) {
                    var cellText = column == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{column + 1}";
                    worksheet.Cells[startRow + row, column].Value = cellText;
                }
            }

            var table = new Table(range.TopRowIndex, range.LeftColumnIndex, range.RowCount, range.ColumnCount);

            if(useCustomTableStyle) {
                table.StyleName = CreateCustomTableStyle(workbook.TableStyles, "CustomTableStyle").Name;
            }
            else {
                table.StyleName = PredefinedTableStyleNames.Dark9;
            }

            if(showTotals) {
                table.HasTotalRow = true;
                table.Columns.Last().TotalRowFunction = TotalRowFunction.Count;
            }

            worksheet.Tables.Add(table, hasHeaders);

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        TableStyle CreateCustomTableStyle(TableStyleCollection tableStyles, string styleName) {
            var customTableStyle = tableStyles.Add(styleName);

            customTableStyle.WholeTable.Format = new DifferentialFormat {
                Font = new DifferentialFontFormat().WithColor(Color.FromArgb(107, 107, 107))
            };

            customTableStyle.HeaderRow.Format = new DifferentialFormat {
                Fill = new DifferentialSolidFillFormat(Color.FromArgb(64, 66, 166)),
                Font = new DifferentialFontFormat().WithColor(Color.White).WithBold(true)
            };

            customTableStyle.TotalRow.Format = new DifferentialFormat {
                Fill = new DifferentialSolidFillFormat(Color.FromArgb(115, 193, 211)),
                Font = new DifferentialFontFormat().WithColor(Color.White).WithBold(true)
            };

            var secondRowStripeStyle = customTableStyle.SecondRowStripe;
            secondRowStripeStyle.Format = new DifferentialFormat {
                Fill = new DifferentialSolidFillFormat(Color.FromArgb(234, 234, 234))
            };
            secondRowStripeStyle.StripeSize = 1;

            return customTableStyle;
        }
        #endregion Code
    }
}
