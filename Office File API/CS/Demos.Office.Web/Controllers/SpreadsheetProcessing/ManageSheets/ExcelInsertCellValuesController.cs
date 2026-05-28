#region Code
using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
#endregion Code

using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelInsertCellValuesController : SpreadsheetControllerBase {
        [HttpPost("insert-cell-values/insert")]
        public async Task<IActionResult> Insert(IFormFile imageFile, [FromForm] ExcelInsertCellValuesModel model) {
            // Read file stream
            using var imageStream = new MemoryStream();
            await imageFile.CopyToAsync(imageStream);

            imageStream.Seek(0, SeekOrigin.Begin);

            // Get values from model
            string textCellValue = model.TextValue;
            double numericCellValue = model.NumericValue;
            DateTime dateCellValue = model.DateValue;
            ExcelNumberFormat numberFormat = model.NumberFormat;
            ExcelNumberFormat dateFormat = model.DateFormat;

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Convert document to the target format
            var convertedStream = InsertCellValues(textCellValue, numericCellValue, dateCellValue, imageStream, numberFormat, dateFormat, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream InsertCellValues(string textCellValue, double numericCellValue, DateTime dateCellValue,
            Stream imageStream, ExcelNumberFormat numberFormat, ExcelNumberFormat dateFormat, DocumentFormat outputFormat) {

            using Workbook workbook = new Workbook();
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            Worksheet firstWorksheet = workbook.Worksheets.First();
            firstWorksheet.Name = "Inserted Cell Values";

            firstWorksheet["A1"].Value = "Text Value";
            firstWorksheet["B1"].Value = "Numeric Value";
            firstWorksheet["C1"].Value = "Date Value";
            firstWorksheet["D1"].Value = "Image";

            // Make header text bold
            var formatting = firstWorksheet["A1:D1"].BeginUpdateFormatting();
            formatting.Font.Bold = true;
            firstWorksheet["A1:D1"].EndUpdateFormatting(formatting);

            // Set cell values
            firstWorksheet["A2"].Value = textCellValue;

            firstWorksheet["B2"].Value = numericCellValue;
            firstWorksheet["B2"].NumberFormat = ExcelCellFormatHelper.GetNumberFormatString(numberFormat);

            firstWorksheet["C2"].Value = dateCellValue;
            firstWorksheet["C2"].NumberFormat = ExcelCellFormatHelper.GetDateFormatString(dateFormat);

            firstWorksheet["D2"].Value = DXImage.FromStream(imageStream);

            firstWorksheet.Columns.AutoFit(0, 4);

            // Set the image column width to 1 inch
            firstWorksheet.Columns["D"].Width = 1;

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
