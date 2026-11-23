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
    public class ExcelRichTextFormattingController : SpreadsheetControllerBase {
        [HttpPost("richtext-formatting/create")]
        public IActionResult CreateWorkbook([FromForm] ExcelRichTextFormattingModel model) {
            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var outputStream = CreateWorkbookWithRichTextFormatting(
                model.FirstTextPart, model.SecondTextPart, model.ThirdTextPart,
                ColorTranslator.FromHtml(model.FirstPartColorHex ?? string.Empty),
                ColorTranslator.FromHtml(model.SecondPartColorHex ?? string.Empty),
                ColorTranslator.FromHtml(model.ThirdPartColorHex ?? string.Empty),
                targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CreateWorkbookWithRichTextFormatting(string firstTextPart, string secondTextPart, string thirdTextPart,
            Color firstPartColor, Color secondPartColor, Color thirdPartColor, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "Rich Text Formatting";

            RichText richText = new RichText()
                .AddRun(firstTextPart + " ", new FontFormat("Arial", 14).WithColor(firstPartColor))
                .AddRun(secondTextPart + " ", new FontFormat("Tahoma", 14).WithColor(secondPartColor))
                .AddRun(thirdTextPart, new FontFormat("Castellar", 14).WithColor(thirdPartColor));

            worksheet.Cells["A1"].SetRichText(richText);

            Stream outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
