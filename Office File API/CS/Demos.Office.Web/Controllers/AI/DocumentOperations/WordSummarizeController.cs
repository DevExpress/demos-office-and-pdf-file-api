#region Code
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations;

namespace BlazorDemo.OfficeFileApi.Controllers.AI.DocumentOperations {
    [ApiController]
    [Route("api/ai")]
    public class WordSummarizeController : WordControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public WordSummarizeController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/word-summarize")]
        public async Task<IActionResult> WordSummarize(IFormFile file, [FromForm] WordSummarizationModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await SummarizeDocument(_aiDocProcessingService, inputStream, model.SummarizationMode, model.OutputMode);

            return CreateOutputFile(outputStream, model.OutputMode == OutputMode.StandaloneFile ? "txt" : "docx", model.OutputName);
        }

        #region Code
        async Task<Stream> SummarizeDocument(IAIDocProcessingService service, Stream inputStream, SummarizationMode summarizationMode, OutputMode outputMode) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load the input document
            wordProcessor.LoadDocument(inputStream);

            var result = await service.SummarizeAsync(wordProcessor, summarizationMode);

            // Create output stream
            var outputStream = new MemoryStream();

            switch(outputMode) {
                case OutputMode.InsertAboveContent:
                    InsertAboveContent(wordProcessor.Document, result);
                    wordProcessor.SaveDocument(outputStream, DocumentFormat.OpenXml);
                    break;
                case OutputMode.InsertBelowContent:
                    InsertBelowContent(wordProcessor.Document, result);
                    wordProcessor.SaveDocument(outputStream, DocumentFormat.OpenXml);
                    break;
                case OutputMode.StandaloneFile:
                    CreateStandaloneFile(wordProcessor, result);
                    wordProcessor.SaveDocument(outputStream, DocumentFormat.PlainText);
                    break;
                default:
                    throw new InvalidOperationException("Invalid output mode");
            }

            outputStream.Position = 0;
            return outputStream;
        }

        void InsertAboveContent(Document document, string text) {
            var section = document.InsertSection(document.Range.Start);
            InsertContentToSection(document, text, section);
        }

        void InsertBelowContent(Document document, string text) {
            document.InsertSection(document.Range.End);
            var section = document.Sections.Last();
            InsertContentToSection(document, text, section);
        }

        void InsertContentToSection(Document document, string text, Section section) {
            // Clear section formatting
            CharacterProperties sectionCharacterProperties = document.BeginUpdateCharacters(section.Range);
            sectionCharacterProperties.Reset(CharacterPropertiesMask.All);
            sectionCharacterProperties.Style.Reset(CharacterPropertiesMask.All);
            document.EndUpdateCharacters(sectionCharacterProperties);
            ParagraphProperties sectionParagraphProperties = document.BeginUpdateParagraphs(section.Range);
            sectionParagraphProperties.Reset(ParagraphPropertiesMask.All);
            sectionParagraphProperties.Style.Reset(ParagraphPropertiesMask.All);
            document.EndUpdateParagraphs(sectionParagraphProperties);

            // Insert summary text
            var textRange = document.InsertText(section.Range.Start, "Document Summary");
            document.Paragraphs.Insert(textRange.End);
            document.Paragraphs.Insert(textRange.End);
            document.InsertText(textRange.End, text);

            // Format summary title
            CharacterProperties titleCharacterProperties = document.BeginUpdateCharacters(section.Paragraphs.First().Range);
            titleCharacterProperties.Bold = true;
            titleCharacterProperties.FontSize = 20;
            document.EndUpdateCharacters(titleCharacterProperties);
        }

        void CreateStandaloneFile(RichEditDocumentServer wordProcessor, string text) {
            wordProcessor.CreateNewDocument();
            wordProcessor.Text = text;
        }
        #endregion Code
    }
}
