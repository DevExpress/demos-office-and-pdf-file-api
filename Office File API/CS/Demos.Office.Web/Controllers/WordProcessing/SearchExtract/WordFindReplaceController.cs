using BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.SearchExtract {
    [ApiController]
    [Route("api/wordprocessing/find-replace")]
    public class WordFindReplaceController : WordControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> ProcessFindReplace([FromForm] IFormFile file, [FromForm] WordFindReplaceModel model) {

            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = new MemoryStream();
            _ = model.Operation switch
 {
                FindReplaceActionType.Replace => outputStream = FindReplace(inputStream, model.SearchText, model.SearchOptions, model.ReplaceText ?? "", targetFormat),
                FindReplaceActionType.Remove => outputStream = FindRemove(inputStream, model.SearchText, model.SearchOptions, targetFormat),
                FindReplaceActionType.Highlight => outputStream = FindHighlight(inputStream, model.SearchText, model.SearchOptions, ColorTranslator.FromHtml(model.HighlightColor), targetFormat),
                _ => throw new InvalidOperationException("Unsupported operation")
            };
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FindReplace(Stream inputStream, string findText, SearchOptions searchOptions, string replaceText, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load document
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            document.ReplaceAll(findText, replaceText, searchOptions);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            return outputStream;
        }

        Stream FindRemove(Stream inputStream, string findText, SearchOptions searchOptions, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load document
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            // Find all occurrences of the search text
            DocumentRange[] searchResults = document.FindAll(findText, searchOptions);

            // Remove found ranges
            for (int i = searchResults.Length - 1; i >= 0; i--)
                document.Delete(searchResults[i]);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            return outputStream;
        }

        Stream FindHighlight(Stream inputStream, string findText, SearchOptions searchOptions, Color highlightColor, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load document
            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            // Find all occurrences of the search text
            DocumentRange[] searchResults = document.FindAll(findText, searchOptions);

            // Apply highlight color to found ranges
            foreach (var range in searchResults) {
                var charProps = document.BeginUpdateCharacters(range);
                charProps.BackColor = highlightColor;
                document.EndUpdateCharacters(charProps);
            }

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            return outputStream;
        }

        #endregion Code
    }
}
