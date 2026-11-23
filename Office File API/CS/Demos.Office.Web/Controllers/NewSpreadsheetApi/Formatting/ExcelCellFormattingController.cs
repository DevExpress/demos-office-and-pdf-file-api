#region Code
using DevExpress.Docs.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.Formatting;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.Formatting {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelCellFormattingController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("cell-formatting/format")]
        public async Task<IActionResult> Format(IFormFile file, [FromForm] ExcelCellFormattingModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var resultingStream = FormatWorkbook(inputStream,
                model.CreateCustomStyle,
                model.FontName, ColorTranslator.FromHtml(model.ForeColorHex ?? string.Empty), model.FontSize,
                model.FontStyle is CellFontStyle.Bold or CellFontStyle.BoldItalic,
                model.FontStyle is CellFontStyle.Italic or CellFontStyle.BoldItalic,
                ColorTranslator.FromHtml(model.BackColorHex ?? string.Empty),
                model.HorizontalAlignment, model.VerticalAlignment,
                targetFormat);

            return CreateOutputFile(resultingStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FormatWorkbook(Stream inputStream, bool createNewStyle, string fontName, Color fontColor, double fontSize,
            bool bold, bool italic, Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment,
            SpreadsheetVerticalAlignment verticalAlignment, DocumentFormat outputFormat) {

            using var workbook = new Workbook(inputStream);

            Worksheet worksheet = workbook.Worksheets[0];
            CellRange dataRange = worksheet.GetDataRange();
            if(dataRange.HasData) {

                var font = new FontFormat(fontName, fontSize, fontColor, bold, italic, false, UnderlineType.None);
                var fill = new SolidFillFormat(backgroundColor);
                var alignment = new AlignmentFormat(horizontalAlignment, verticalAlignment);

                if(createNewStyle) {
                    Style customStyle = CreateCustomStyle(workbook, font, fill, alignment);
                    ApplyCustomStyle(dataRange, customStyle);
                }
                else {
                    ApplyDirectFormatting(dataRange, font, fill, alignment);
                }

                dataRange.AutoFitColumns();
            }

            Stream outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }

        Style CreateCustomStyle(Workbook workbook, FontFormat font, FillFormat fill, AlignmentFormat alignment) {
            var customStyle = new Style("Custom Style") {
                Font = font,
                Fill = fill,
                Alignment = alignment
            };

            workbook.Styles.Add(customStyle);
            return customStyle;
        }

        void ApplyCustomStyle(CellRange cellRange, Style style) {
            cellRange.ApplyFont(null);
            cellRange.ApplyFill(null);
            cellRange.ApplyAlignment(null);

            cellRange.StyleName = style.Name;
        }

        void ApplyDirectFormatting(CellRange cellRange, FontFormat font, FillFormat fill, AlignmentFormat alignment) {
            cellRange.ApplyFont(font);
            cellRange.ApplyFill(fill);
            cellRange.ApplyAlignment(alignment);
        }
        #endregion Code
    }
}
