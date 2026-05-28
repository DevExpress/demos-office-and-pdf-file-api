using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/footnotes-endnotes")]
    public class WordFootnotesEndnotesController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateFileStream = new FileStream(DemoDocumentPaths.WordNoteTemplateFile, FileMode.Open, FileAccess.Read);
            var templateStream = new MemoryStream();
            templateFileStream.CopyTo(templateStream);
            templateStream.Position = 0;

            return File(templateStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "FootnoteEndnote_Template.docx");
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessFootnotesEndnotes([FromForm] WordFootnotesEndnotesModel model) {
            using var templateFileStream = new FileStream(DemoDocumentPaths.WordNoteTemplateFile, FileMode.Open, FileAccess.Read);
            var templateStream = new MemoryStream();
            templateFileStream.CopyTo(templateStream);
            templateStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process footnotes/endnotes document
            var outputStream = await ProcessFootnotesEndnotesDocument(templateStream, model, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        async Task<Stream> ProcessFootnotesEndnotesDocument(Stream templateStream, WordFootnotesEndnotesModel model, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(templateStream);
            var document = wordProcessor.Document;
            // Add footnotes if requested
            if (model.AddNewFootnote)
                AddNewFootnote(document, document.Bookmarks["footnote"].Range.Start, "This is a new footnote");

            // Add endnotes if requested
            if (model.AddNewEndnote)
                AddNewEndnote(document, document.Bookmarks["endnote"].Range.Start, "This is a new endnote");

            ChangeNoteAppearance(document, model.NoteSeparator, model.NoteRestartType, model.NumberingFormat, model.StartNumber);
            ChangeNotePosition(document, model.EndnotePosition, model.FootnotePosition);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #region Code
        void AddNewFootnote(Document document, DocumentPosition footnotePosition, string footnoteText) {
            Note footnote = document.Footnotes.Insert(footnotePosition);
            SubDocument footnoteDocument = footnote.BeginUpdate();
            var footnoteRange = footnoteDocument.AppendText(footnoteText);
            CharacterProperties footnoteCharProps = footnoteDocument.BeginUpdateCharacters(footnoteRange);
            footnoteCharProps.Bold = true;
            footnoteCharProps.ForeColor = Color.Red;
            footnoteDocument.EndUpdateCharacters(footnoteCharProps);
            footnote.EndUpdate(footnoteDocument);
        }

        void AddNewEndnote(Document document, DocumentPosition endnotePosition, string endnoteText) {
            Note endnote = document.Endnotes.Insert(endnotePosition);
            SubDocument endnoteDocument = endnote.BeginUpdate();
            var endnoteRange = endnoteDocument.AppendText(endnoteText);
            CharacterProperties endnoteCharProps = endnoteDocument.BeginUpdateCharacters(endnoteRange);
            endnoteCharProps.Bold = true;
            endnoteCharProps.ForeColor = Color.Red;
            endnoteDocument.EndUpdateCharacters(endnoteCharProps);
            endnote.EndUpdate(endnoteDocument);
        }

        void ChangeNoteAppearance(Document document, string noteSeparator, NoteRestartType restartType,
            NumberingFormat format, int startNumber) {
            SubDocument noteSeparatorDoc = document.Footnotes.BeginUpdateSeparator(NoteSeparatorType.Separator);
            noteSeparatorDoc.Delete(noteSeparatorDoc.Range);
            noteSeparatorDoc.AppendText(noteSeparator);
            document.Footnotes.EndUpdateSeparator(noteSeparatorDoc);

            foreach (Section section in document.Sections) {
                FootnoteOptions footnoteOptions = section.FootnoteOptions;
                footnoteOptions.RestartType = restartType;
                footnoteOptions.NumberingFormat = format;
                footnoteOptions.StartNumber = startNumber;
            }
        }

        void ChangeNotePosition(Document document, EndnotePosition endnotePosition, FootnotePosition footnotePosition) {
            document.EndnotePosition = endnotePosition;

            foreach (Section section in document.Sections)
                section.FootnoteOptions.Position = footnotePosition;
        }
        #endregion Code
    }
}
