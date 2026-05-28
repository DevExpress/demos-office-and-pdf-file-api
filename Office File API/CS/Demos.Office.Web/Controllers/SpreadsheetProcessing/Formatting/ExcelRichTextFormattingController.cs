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
    public class ExcelRichTextFormattingController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("richtext-formatting/create")]
        public async Task<IActionResult> CreateWorkbook([FromForm] ExcelRichTextFormattingModel model) {
            // Create input stream
            using var inputStream = new MemoryStream();

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate a workbook with Rich Text Formatting
            var outputStream = await CreateWorkbookWithRichTextFormatting(
                inputStream,
                model.FirstTextPart, model.SecondTextPart, model.ThirdTextPart,
                ColorTranslator.FromHtml(model.FirstPartColorHex ?? string.Empty),
                ColorTranslator.FromHtml(model.SecondPartColorHex ?? string.Empty),
                ColorTranslator.FromHtml(model.ThirdPartColorHex ?? string.Empty),
                targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        async Task<Stream> CreateWorkbookWithRichTextFormatting(Stream inputStream, string firstTextPart, string secondTextPart,
            string thirdTextPart, Color firstPartColor, Color secondPartColor, Color thirdPartColor, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            workbook.BeginUpdate();
            try {
                // Use the first worksheet
                Worksheet worksheet = workbook.Worksheets.First();
                worksheet.Name = "Rich Text Formatting";

                // Create a RichTextString instance
                RichTextString richText = new RichTextString();

                // Add three text runs. Each run has its own font settings
                richText.AddTextRun(firstTextPart + " ", new RichTextRunFont("Arial", 14, firstPartColor));
                richText.AddTextRun(secondTextPart + " ", new RichTextRunFont("Tahoma", 14, secondPartColor));
                richText.AddTextRun(thirdTextPart, new RichTextRunFont("Castellar", 14, thirdPartColor));

                // Assign the rich formatted text to the cell A1
                worksheet["A1"].SetRichText(richText);
            }
            finally {
                workbook.EndUpdate();
            }

            // Create output stream
            Stream outputStream = new MemoryStream();

            // Save document in target format
            await workbook.SaveDocumentAsync(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
