#region Code
using DevExpress.XtraRichEdit;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentConversion;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentConversion {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordToWordController : WordControllerBase {
        [HttpPost("word-to-word/convert")]
        public async Task<IActionResult> ConvertWordToWord(IFormFile file, [FromForm] WordToWordConversionModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            // Convert document
            var convertedStream = ConvertDocument(inputStream, model.EmbedFonts, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ConvertDocument(Stream inputStream, bool embedFonts, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            // Configure HTML export settings for image handling
            wordProcessor.Options.Export.Html.EmbedImages = true;
            wordProcessor.Options.Export.Mht.EmbedImages = true;

            // Load the input document
            wordProcessor.LoadDocument(inputStream);

            wordProcessor.Document.EmbedFonts = embedFonts;

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            wordProcessor.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
