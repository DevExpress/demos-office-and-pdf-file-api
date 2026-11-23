#region Code

using DevExpress.Docs.Word;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentConversion;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentConversion {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordToWordController : WordControllerBase {
        [HttpPost("word-to-word/convert")]
        public async Task<IActionResult> ConvertWordToWord(IFormFile file, [FromForm] WordToWordConversionModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            // Convert document
            var convertedStream = ConvertDocument(inputStream, model.EmbedFonts, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ConvertDocument(Stream inputStream, bool embedFonts, DocumentFormat outputFormat) {
            // Load the input document
            using var document = new WordDocument(inputStream);

            document.EmbedFonts = embedFonts;

            var saveOptions = new SaveOptions() {
                DocumentFormat = outputFormat
            };
            // Configure HTML export settings for image handling
            saveOptions.HtmlOptions.EmbedImages = true;
            saveOptions.MhtOptions.EmbedImages = true;

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            document.Save(outputStream, saveOptions);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}