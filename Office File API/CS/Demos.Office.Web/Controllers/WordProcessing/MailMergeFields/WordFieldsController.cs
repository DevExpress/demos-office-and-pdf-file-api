using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.MailMergeFields;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.MailMergeFields {
    [ApiController]
    [Route("api/wordprocessing/fields")]
    public class WordFieldsController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordFieldTemplateFile, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "MailMerge_Template.docx");
        }

        [HttpPost("process")]
        public async Task<IActionResult> CreateMailMergedDocument([FromForm] WordFieldsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordFieldTemplateFile, FileMode.OpenOrCreate);
            Stream stream = new MemoryStream();
            templateStream.CopyTo(stream);
            stream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process track changes on the template
            var outputStream = model.Action switch {
                FieldAction.Insert => InsertFields(stream, model.FieldEntryValues, targetFormat),
                FieldAction.Update => UpdateFields(stream, targetFormat),
                FieldAction.Lock => LockFields(stream, model.FieldEntryValues, targetFormat),
                FieldAction.Unlink => UnlinkFields(stream, model.FieldEntryValues, targetFormat),
                FieldAction.Delete => DeleteFields(stream, model.FieldEntryValues, targetFormat),
                _ => null,
            };

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        MemoryStream InsertFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(stream);
            Document document = wordProcessor.Document;
            Table table = document.Tables.First();
            document.CalculateDocumentVariable += CalculateDocumentVariable;
            foreach (var fieldType in fieldNames) {
                var row = table.Rows.Append();
                var firstCell = row.FirstCell;
                var secondCell = row.LastCell;
                switch (fieldType) {
                    case "DATE":
                        document.InsertText(firstCell.Range.Start, "New Date Field");
                        document.Fields.Create(secondCell.Range.Start, "DATE");
                        break;
                    case "TIME":
                        document.InsertText(firstCell.Range.Start, "New Time Field");
                        document.Fields.Create(secondCell.Range.Start, "TIME");
                        break;
                    case "PAGE":
                        document.InsertText(firstCell.Range.Start, "New Page Field");
                        document.Fields.Create(secondCell.Range.Start, "PAGE");
                        break;
                    case "NUMPAGES":
                        document.InsertText(firstCell.Range.Start, "New NumPages Field");
                        document.Fields.Create(secondCell.Range.Start, "NUMPAGES");
                        break;
                    case "DocVariable":
                        document.InsertText(firstCell.Range.Start, "New DocVariable Field");
                        document.Fields.Create(secondCell.Range.Start, "DocVariable MyVariable");
                        break;
                    case "AUTHOR":
                        document.InsertText(firstCell.Range.Start, "New Author Field");
                        document.DocumentProperties.Author = "Test Author";
                        document.Fields.Create(secondCell.Range.Start, "AUTHOR");
                        break;
                    case "TITLE":
                        document.InsertText(firstCell.Range.Start, "New Title Field");
                        document.DocumentProperties.Title = "Test Title";
                        document.Fields.Create(secondCell.Range.Start, "TITLE");
                        break;
                    case "DOCPROPERTY":
                        document.InsertText(firstCell.Range.Start, "New DocProperty Field");
                        document.CustomProperties.Add("MyProperty", "Custom Property Value");
                        document.Fields.Create(secondCell.Range.Start, "DOCPROPERTY MyProperty");
                        break;
                }
            }
            document.UpdateAllFields();
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        MemoryStream UpdateFields(Stream stream, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.CalculateDocumentVariable += CalculateDocumentVariable;
            wordProcessor.LoadDocument(stream);
            Document document = wordProcessor.Document;
            document.UpdateAllFields();
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        MemoryStream LockFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(stream);
            Document document = wordProcessor.Document;
            document.CalculateDocumentVariable += CalculateDocumentVariable;
            document.UpdateAllFields();
            foreach (Field field in document.Fields) {
                field.ShowCodes = true;
                string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
                if (fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                    field.Locked = true;
            }
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        MemoryStream UnlinkFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(stream);
            Document document = wordProcessor.Document;
            document.CalculateDocumentVariable += CalculateDocumentVariable;
            document.UpdateAllFields();
            for (int i = document.Fields.Count - 1; i >= 0; i--) {
                Field field = document.Fields[i];
                field.ShowCodes = true;
                string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
                if (fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                    field.Unlink();
            }
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        MemoryStream DeleteFields(Stream stream, IEnumerable<string> fieldNames, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(stream);
            Document document = wordProcessor.Document;
            document.CalculateDocumentVariable += CalculateDocumentVariable;
            document.UpdateAllFields();
            for (int i = document.Fields.Count - 1; i >= 0; i--) {
                Field field = document.Fields[i];
                field.ShowCodes = true;
                string fieldCode = document.GetText(field.CodeRange).ToUpper().TrimStart();
                if (fieldNames.Any(fn => fieldCode.StartsWith(fn, StringComparison.InvariantCultureIgnoreCase)))
                    document.Delete(field.Range);
            }
            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        void CalculateDocumentVariable(object sender, CalculateDocumentVariableEventArgs e) {
            if (e.VariableName == "MyVariable") {
                e.Value = "My Test Variable " + DateTime.Now.ToString("MMMM dd, yyyy");
                e.Handled = true;
            }
            if (e.VariableName == "Test") {
                var tempWordProcessor = new RichEditDocumentServer();
                tempWordProcessor.Unit = DevExpress.Office.DocumentUnit.Inch;
                Document tempDoc = tempWordProcessor.Document;
                Shape picture = tempDoc.Shapes.InsertPicture(tempDoc.Range.Start, DocumentImageSource.FromFile(DemoDocumentPaths.DevExpressLogo));
                picture.TextWrapping = TextWrappingType.InLineWithText;
                picture.LockAspectRatio = true;
                picture.Width = 2f;
                e.Value = tempWordProcessor;
                e.Handled = true;
            }
        }
        #endregion Code
    }
}
