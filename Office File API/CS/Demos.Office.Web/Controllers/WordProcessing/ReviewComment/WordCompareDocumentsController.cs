using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.ReviewComment;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ReviewComment {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordCompareDocumentsController : WordControllerBase {
        [HttpPost("compare-documents/compare")]
        public async Task<IActionResult> CompareDocuments(IFormFile file1, IFormFile file2, [FromForm] WordCompareDocumentsModel model) {
            // Load original document from stream
            using var originalStream = new MemoryStream();
            await file1.CopyToAsync(originalStream);
            originalStream.Position = 0;

            // Load revised document from stream
            using var revisedStream = new MemoryStream();
            await file2.CopyToAsync(revisedStream);
            revisedStream.Position = 0;

            // Compare documents
            var comparedStream = CompareDocuments(originalStream, revisedStream, model.Author,
                model.CaseChanges, model.Formatting, model.HeadersFooters,
                model.TextBoxes, model.ComparisonLevel, WordFormatHelper.GetDocumentFormat(model.OutputFormat));

            return CreateOutputFile(comparedStream, model.OutputFormat, model.OutputName);

        }

        #region Code
        Stream CompareDocuments(Stream originalStream, Stream revisedStream, string author,
            bool compareCaseChanges, bool compareFormatting, bool compareHeadersFooters, bool compareTextBoxes,
            ComparisonLevel comparisonLevel, DocumentFormat outputFormat) {
            // Create document servers for comparison
            using var originalWordProcessor = new RichEditDocumentServer();
            using var revisedWordProcessor = new RichEditDocumentServer();

            // Load documents
            originalWordProcessor.LoadDocument(originalStream);
            revisedWordProcessor.LoadDocument(revisedStream);

            // Get documents for comparison
            var originalDoc = originalWordProcessor.Document;
            var revisedDoc = revisedWordProcessor.Document;

            CompareDocumentOptions options = new CompareDocumentOptions();
            options.CompareFormatting = compareFormatting;
            options.CompareCaseChanges = compareCaseChanges;
            options.CompareHeadersAndFooters = compareHeadersFooters;
            options.CompareTextBoxes = compareTextBoxes;
            options.ComparisonLevel = comparisonLevel;
            options.Author = author;
            options.DateTime = DateTime.Now;

            Document comparedDocument = originalDoc.Compare(revisedDoc, options);

            // Save the result document to a stream
            var outputStream = new MemoryStream();
            comparedDocument.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion Code
    }
}
