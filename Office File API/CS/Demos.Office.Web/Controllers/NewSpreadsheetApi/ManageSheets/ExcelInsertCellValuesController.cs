#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Drawing;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.ManageSheets {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelInsertCellValuesController : SpreadsheetControllerBase {
        [HttpPost("insert-cell-values/insert")]
        public async Task<IActionResult> Insert(IFormFile imageFile, [FromForm] ExcelInsertCellValuesModel model) {
            using var imageStream = new MemoryStream();
            await imageFile.CopyToAsync(imageStream);
            imageStream.Position = 0;

            string textCellValue = model.TextValue;
            double numericCellValue = model.NumericValue;
            DateTime dateCellValue = model.DateValue;
            ExcelNumberFormat numberFormat = model.NumberFormat;
            ExcelNumberFormat dateFormat = model.DateFormat;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = InsertCellValues(textCellValue, numericCellValue, dateCellValue, imageStream, numberFormat, dateFormat, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream InsertCellValues(string textCellValue, double numericCellValue, DateTime dateCellValue,
            Stream imageStream, ExcelNumberFormat numberFormat, ExcelNumberFormat dateFormat, DocumentFormat outputFormat) {

            using var workbook = new Workbook();

            Worksheet firstWorksheet = workbook.Worksheets[0];
            firstWorksheet.Name = "Inserted Cell Values";

            firstWorksheet.Cells["A1"].Value = "Text Value";
            firstWorksheet.Cells["B1"].Value = "Numeric Value";
            firstWorksheet.Cells["C1"].Value = "Date Value";
            firstWorksheet.Cells["D1"].Value = "Image";

            firstWorksheet["A1:D1"].ApplyFont(new FontFormat(FontFormat.DefaultName, FontFormat.DefaultSize).WithBold(true));

            firstWorksheet.Cells["A2"].Value = textCellValue;

            firstWorksheet.Cells["B2"].Value = numericCellValue;
            firstWorksheet.Cells["B2"].Format.NumberFormat = ExcelCellFormatHelper.GetNumberFormatString(numberFormat);

            firstWorksheet.Cells["C2"].SetDateValue(dateCellValue, ExcelCellFormatHelper.GetDateFormatString(dateFormat));

            firstWorksheet["A1:C2"].AutoFitColumns();

            firstWorksheet.Cells["D2"].SetImage(DXImage.FromStream(imageStream));

            firstWorksheet.Columns["D"].Width = 18;

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
