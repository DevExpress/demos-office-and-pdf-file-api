#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Formatting;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Formatting {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelCellFormattingController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("cell-formatting/format")]
        public async Task<IActionResult> Format(IFormFile file, [FromForm] ExcelCellFormattingModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Format workbook
            var resultingStream = FormatWorkbook(inputStream,
                model.CreateCustomStyle,
                model.FontName, ColorTranslator.FromHtml(model.ForeColorHex ?? string.Empty), model.FontSize, model.FontStyle,
                ColorTranslator.FromHtml(model.BackColorHex ?? string.Empty),
                model.HorizontalAlignment, model.VerticalAlignment,
                targetFormat);

            return CreateOutputFile(resultingStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FormatWorkbook(Stream inputStream, bool createNewStyle, string fontName, Color fontColor, double fontSize,
            SpreadsheetFontStyle fontStyle, Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment,
            SpreadsheetVerticalAlignment verticalAlignment, DocumentFormat outputFormat) {

            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            //Get the data range of the active worksheet
            CellRange dataRange = workbook.Worksheets.ActiveWorksheet.GetDataRange();

            if(createNewStyle) {
                Style customStyle = CreateCustomStyle(workbook, fontName, fontColor, fontSize, fontStyle, backgroundColor, horizontalAlignment, verticalAlignment);
                dataRange.Style = customStyle;
            }
            else {
                ApplyDirectFormatting(dataRange, fontName, fontColor, fontSize, fontStyle, backgroundColor, horizontalAlignment, verticalAlignment);
            }

            dataRange.AutoFitColumns();

            // Create output stream
            Stream outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }

        Style CreateCustomStyle(Workbook workbook, string fontName, Color fontColor, double fontSize, SpreadsheetFontStyle fontStyle,
            Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment, SpreadsheetVerticalAlignment verticalAlignment) {

            // Add a new style under the "Custom Style" name to the style collection
            Style customStyle = workbook.Styles.Add("Custom Style");

            // Specify the style's format characteristics.
            customStyle.BeginUpdate();
            try {
                // Specify font settings (font name, color, size and style)
                customStyle.Font.Name = fontName;
                customStyle.Font.Color = fontColor;
                customStyle.Font.Size = fontSize;
                customStyle.Font.FontStyle = fontStyle;

                // Set the background fill
                customStyle.Fill.BackgroundColor = backgroundColor;

                // Specify text alignment in cells
                customStyle.Alignment.Horizontal = horizontalAlignment;
                customStyle.Alignment.Vertical = verticalAlignment;
            }
            finally {
                customStyle.EndUpdate();
            }
            return customStyle;
        }

        void ApplyDirectFormatting(CellRange cellRange, string fontName, Color fontColor, double fontSize, SpreadsheetFontStyle fontStyle,
            Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment, SpreadsheetVerticalAlignment verticalAlignment) {

            // Begin updating of the range formatting
            DevExpress.Spreadsheet.Formatting formatting = cellRange.BeginUpdateFormatting();
            try {
                // Specify font settings (font name, color, size and style)              
                formatting.Font.Name = fontName;
                formatting.Font.Color = fontColor;
                formatting.Font.Size = fontSize;
                formatting.Font.FontStyle = fontStyle;

                // Specify cell background color            
                formatting.Fill.BackgroundColor = backgroundColor;

                // Specify text alignment in cells
                formatting.Alignment.Horizontal = horizontalAlignment;
                formatting.Alignment.Vertical = verticalAlignment;
            }
            finally {
                cellRange.EndUpdateFormatting(formatting);
            }
        }
        #endregion Code
    }
}
