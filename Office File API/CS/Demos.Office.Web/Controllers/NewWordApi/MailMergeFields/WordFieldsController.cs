using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.MailMergeFields;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
using DevExpress.Drawing;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.MailMergeFields {
    [ApiController]
    [Route("api/newwordapi/fields")]
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
        public IActionResult CreateMailMergedDocument([FromForm] WordFieldsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordFieldTemplateFile, FileMode.Open, FileAccess.Read);
            var stream = new MemoryStream();
            templateStream.CopyTo(stream);
            stream.Position = 0;

            var fieldNames = model.FieldEntryValues ?? Array.Empty<string>();
            // Process track changes on the template
            var outputStream = model.Action switch {
                FieldAction.Insert => InsertFields(stream, fieldNames, model.OutputFormat),
                FieldAction.Update => UpdateFields(stream, model.OutputFormat),
                FieldAction.Lock => LockFields(stream, fieldNames, model.OutputFormat),
                FieldAction.Unlink => UnlinkFields(stream, fieldNames, model.OutputFormat),
                FieldAction.Delete => DeleteFields(stream, fieldNames, model.OutputFormat),
                _ => null,
            };

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        MemoryStream InsertFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
            using var document = new WordDocument(stream);
            var table = document.Body.OfType<Table>().First();
            foreach(var fieldType in fieldNames) {
                switch(fieldType) {
                    case "DATE":
                        AppendFieldRow(table, "New Date Field", "DATE");
                        break;
                    case "TIME":
                        AppendFieldRow(table, "New Time Field", "TIME");
                        break;
                    case "PAGE":
                        AppendFieldRow(table, "New Page Field", "PAGE");
                        break;
                    case "NUMPAGES":
                        AppendFieldRow(table, "New NumPages Field", "NUMPAGES");
                        break;
                    case "DocVariable":
                        AppendFieldRow(table, "New DocVariable Field", "DocVariable MyVariable");
                        break;
                    case "AUTHOR":
                        document.DocumentProperties.Author = "Test Author";
                        AppendFieldRow(table, "New Author Field", "AUTHOR");
                        break;
                    case "TITLE":
                        document.DocumentProperties.Title = "Test Title";
                        AppendFieldRow(table, "New Title Field", "TITLE");
                        break;
                    case "DOCPROPERTY":
                        document.DocumentProperties.CustomProperties["MyProperty"] = "Custom Property Value";
                        AppendFieldRow(table, "New DocProperty Field", "DOCPROPERTY MyProperty");
                        break;
                }
            }

            UpdateAllFields(document);
            return SaveDocument(document, outputFormat);
        }

        void AppendFieldRow(Table table, string caption, string fieldCode) {
            var row = table.Rows.Append();
            row.Cells.First.Content.Add(new Paragraph(caption));
            var fieldParagraph = new Paragraph();
            fieldParagraph.InlineElements.Add(new Field(fieldCode));
            row.Cells.Last.Content.Add(fieldParagraph);
        }

        MemoryStream UpdateFields(Stream stream, string outputFormat) {
            using var document = new WordDocument(stream);
            UpdateAllFields(document);
            return SaveDocument(document, outputFormat);
        }

        MemoryStream LockFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
            using var document = new WordDocument(stream);
            UpdateAllFields(document);
            foreach(var paragraph in GetParagraphs(document.Body)) {
                foreach(var field in paragraph.InlineElements.OfType<Field>()) {
                    if(HasFieldType(field, fieldNames))
                        field.Locked = true;
                }
            }

            return SaveDocument(document, outputFormat);
        }

        MemoryStream UnlinkFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
            using var document = new WordDocument(stream);
            UpdateAllFields(document);
            foreach(var paragraph in GetParagraphs(document.Body)) {
                var inlineElements = paragraph.InlineElements;
                for(int i = inlineElements.Count - 1; i >= 0; i--) {
                    if(inlineElements[i] is Field field && HasFieldType(field, fieldNames)) {
                        var resultElements = field.Result.OfType<Inline>().Select(element => (Inline)element.Clone()).ToList();
                        inlineElements.RemoveAt(i);
                        for(int j = 0; j < resultElements.Count; j++)
                            inlineElements.Insert(i + j, resultElements[j]);
                    }
                }
            }

            return SaveDocument(document, outputFormat);
        }

        MemoryStream DeleteFields(Stream stream, IEnumerable<string> fieldNames, string outputFormat) {
            using var document = new WordDocument(stream);
            UpdateAllFields(document);
            foreach(var paragraph in GetParagraphs(document.Body)) {
                var inlineElements = paragraph.InlineElements;
                for(int i = inlineElements.Count - 1; i >= 0; i--) {
                    if(inlineElements[i] is Field field && HasFieldType(field, fieldNames))
                        inlineElements.RemoveAt(i);
                }
            }

            return SaveDocument(document, outputFormat);
        }

        List<Paragraph> GetParagraphs(IEnumerable<Block> blocks) {
            var result = new List<Paragraph>();
            foreach(var block in blocks) {
                if(block is Paragraph paragraph)
                    result.Add(paragraph);
                else if(block is Table table) {
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells)
                            result.AddRange(GetParagraphs(cell.Content));
                    }
                }
            }

            return result;
        }

        bool HasFieldType(Field field, IEnumerable<string> fieldNames) {
            var fieldCode = field.Code.GetText().TrimStart();
            return fieldNames.Any(fieldName => fieldCode.StartsWith(fieldName, StringComparison.InvariantCultureIgnoreCase));
        }

        void UpdateAllFields(WordDocument document) {
            document.UpdateAllFields(new FieldUpdateOptions { DocumentVariableProvider = new DemoDocumentVariableProvider() });
        }

        MemoryStream SaveDocument(WordDocument document, string outputFormat) {
            var outputStream = new MemoryStream();
            document.Save(outputStream, NewWordFormatHelper.GetDocumentFormat(outputFormat));
            outputStream.Position = 0;
            return outputStream;
        }

        class DemoDocumentVariableProvider : IDocumentVariableProvider {
            const float LogoWidth = 600f;

            public DocumentVariableValue GetValue(string variableName, IList<string> arguments) {
                if(variableName == "MyVariable")
                    return new DocumentVariableValue("My Test Variable " + DateTime.Now.ToString("MMMM dd, yyyy"));
                if(variableName == "Test")
                    return new DocumentVariableValue(CreateLogoDocument());
                return null;
            }

            WordDocument CreateLogoDocument() {
                var image = DXImage.FromStream(new MemoryStream(System.IO.File.ReadAllBytes(DemoDocumentPaths.DevExpressLogo)));
                var picture = new PictureShape(image) {
                    LockAspectRatio = true,
                    Width = LogoWidth,
                    Height = LogoWidth * image.Height / image.Width
                };
                var result = new WordDocument();
                result.Body.Add(picture);
                return result;
            }
        }
        #endregion Code
    }
}