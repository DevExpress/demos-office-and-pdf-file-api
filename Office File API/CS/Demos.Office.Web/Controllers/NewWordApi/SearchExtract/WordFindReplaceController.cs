using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.SearchExtract;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.SearchExtract {
    [ApiController]
    [Route("api/newwordapi/find-replace")]
    public class WordFindReplaceController : WordControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> ProcessFindReplace([FromForm] IFormFile file, [FromForm] WordFindReplaceModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var outputStream = model.Operation switch {
                FindReplaceActionType.Replace => FindReplace(inputStream, model.SearchText, model.SearchOptions, model.ReplaceText ?? "", targetFormat),
                FindReplaceActionType.Remove => FindRemove(inputStream, model.SearchText, model.SearchOptions, targetFormat),
                _ => throw new InvalidOperationException("Unsupported operation")
            };
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream FindReplace(Stream inputStream, string findText, TextSearchOptions searchOptions, string replaceText, DocumentFormat outputFormat) {
            // Load document
            using var document = new WordDocument(inputStream);

            document.Body.ReplaceText(findText, replaceText, searchOptions);

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            return outputStream;
        }

        Stream FindRemove(Stream inputStream, string findText, TextSearchOptions searchOptions, DocumentFormat outputFormat) {
            // Load document
            using var document = new WordDocument(inputStream);

            // Find all occurrences of the search text
            var searchResults = document.Body.FindText(findText, searchOptions);

            // Remove found ranges
            for(int i = searchResults.Count - 1; i >= 0; i--)
                document.Body.Delete(searchResults[i]);

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            return outputStream;
        }
        #endregion Code
    }
}