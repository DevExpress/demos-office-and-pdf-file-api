using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs;
using DevExpress.Docs.Word;
using System.Drawing;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.Formatting {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordTableStyleController : WordControllerBase {
        [HttpGet("table-style/template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordAppointmentFormTemplateFile, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "TableStyle_Template.docx");
        }

        [HttpPost("table-style/create")]
        public IActionResult CreateWordTableStyle([FromForm] WordTableStyleModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordAppointmentFormTemplateFile, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            // Process track changes on the template
            var outputStream = CreateApplyTableStyle(memoryStream, (TableLookTypes)model.TableLookValues, model.OutputFormat);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateApplyTableStyle(Stream stream, TableLookTypes tableLookTypes, string outputFormat) {
            using var document = new WordDocument(stream);
            var table = document.Body.OfType<Table>().First();

            var tableStyle = CreateTableStyle("CustomTableStyle");
            // Add the style to the document collection
            document.TableStyles.Add(tableStyle);
            table.StyleName = tableStyle.Name;
            table.Look = tableLookTypes;

            var outputStream = new MemoryStream();
            document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(outputFormat));
            outputStream.Position = 0;
            return outputStream;
        }

        TableStyle CreateTableStyle(string styleName) {
            // Create a new table style
            var tableStyle = new TableStyle(styleName);

            // Specify style options
            var borders = tableStyle.TableProperties.Borders;
            var innerBorderColor = new OfficeColor(Color.White);
            borders.InsideHorizontal = new TableBorder(BorderLineStyle.Single, 0, innerBorderColor);
            borders.InsideVertical = new TableBorder(BorderLineStyle.Single, 0, innerBorderColor);

            tableStyle.TableCellProperties.BackColor = new OfficeColor(Color.FromArgb(227, 238, 220));

            var conditionalStyles = tableStyle.ConditionalStyles;
            conditionalStyles.Add(TableConditionalStyleType.FirstRow, CreateConditionalStyle(Color.FromArgb(173, 209, 158)));
            conditionalStyles.Add(TableConditionalStyleType.LastRow, CreateConditionalStyle(Color.FromArgb(195, 223, 179)));
            conditionalStyles.Add(TableConditionalStyleType.FirstColumn, CreateConditionalStyle(Color.FromArgb(210, 228, 200)));
            conditionalStyles.Add(TableConditionalStyleType.LastColumn, CreateConditionalStyle(Color.FromArgb(210, 228, 200)));
            conditionalStyles.Add(TableConditionalStyleType.OddRowBanding, CreateConditionalStyle(Color.FromArgb(235, 243, 229)));
            conditionalStyles.Add(TableConditionalStyleType.EvenRowBanding, CreateConditionalStyle(Color.FromArgb(255, 255, 255)));
            conditionalStyles.Add(TableConditionalStyleType.OddColumnBanding, CreateConditionalStyle(Color.FromArgb(239, 247, 233)));
            conditionalStyles.Add(TableConditionalStyleType.EvenColumnBanding, CreateConditionalStyle(Color.FromArgb(227, 238, 220)));
            return tableStyle;
        }

        TableConditionalStyle CreateConditionalStyle(Color cellBackgroundColor) {
            var result = new TableConditionalStyle();
            result.TableCellProperties.BackColor = new OfficeColor(cellBackgroundColor);
            return result;
        }
        #endregion Code
    }
}