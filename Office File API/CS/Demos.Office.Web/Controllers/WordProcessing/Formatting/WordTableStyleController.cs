using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.Formatting;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.Formatting {
    [ApiController]
    [Route("api/wordprocessing/table-style")]
    public class WordTableStyleController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordAppointmentFormTemplateFile, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "TableStyle_Template.docx");
        }

        [HttpPost("create")]
        public async Task<IActionResult> CreateWordTableStyle([FromForm] WordTableStyleModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordAppointmentFormTemplateFile, FileMode.OpenOrCreate);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process track changes on the template
            var outputStream = CreateApplyTableStyle(memoryStream, (TableLookTypes)model.TableLookValues, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CreateApplyTableStyle(Stream stream, TableLookTypes tableLookTypes, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(stream);
            var document = wordProcessor.Document;
            Table table = document.Tables[0];

            table.Style = CreateTableStyle(document, "CustomTableStyle");
            table.TableLook = tableLookTypes;
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        TableStyle CreateTableStyle(Document document, string styleName) {
            // Create a new table style
            var tableStyle = document.TableStyles.CreateNew();
            tableStyle.Name = styleName;

            // Specify style options
            TableBorders tableBorders = tableStyle.TableBorders;
            TableBorder innerHborder = tableBorders.InsideHorizontalBorder;
            innerHborder.LineStyle = BorderLineStyle.Single;
            innerHborder.LineColor = Color.White;

            TableBorder innerVborder = tableBorders.InsideVerticalBorder;
            innerHborder.LineStyle = BorderLineStyle.Single;
            innerVborder.LineColor = Color.White;

            tableStyle.CellBackgroundColor = Color.FromArgb(227, 238, 220);

            TableConditionalStyleProperties conditionalStyleProperties = tableStyle.ConditionalStyleProperties;

            TableConditionalStyle styleForFirstRow = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow);
            styleForFirstRow.CellBackgroundColor = Color.FromArgb(173, 209, 158);
            TableConditionalStyle styleForLastRow = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.LastRow);
            styleForLastRow.CellBackgroundColor = Color.FromArgb(195, 223, 179);

            TableConditionalStyle styleForFirstColumn = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn);
            styleForFirstColumn.CellBackgroundColor = Color.FromArgb(210, 228, 200);
            TableConditionalStyle styleForLastColumn = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.LastColumn);
            styleForLastColumn.CellBackgroundColor = Color.FromArgb(210, 228, 200);

            TableConditionalStyle styleForOddRows = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddRowBanding);
            styleForOddRows.CellBackgroundColor = Color.FromArgb(235, 243, 229);
            TableConditionalStyle styleForEvenRows = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenRowBanding);
            styleForEvenRows.CellBackgroundColor = Color.FromArgb(255, 255, 255);

            TableConditionalStyle styleForOddColumns = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding);
            styleForOddColumns.CellBackgroundColor = Color.FromArgb(239, 247, 233);
            TableConditionalStyle styleForEvenColumns = conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding);
            styleForEvenColumns.CellBackgroundColor = Color.FromArgb(227, 238, 220);

            // Add the style to the document collection
            document.TableStyles.Add(tableStyle);
            return tableStyle;
        }

        #endregion Code
    }
}
