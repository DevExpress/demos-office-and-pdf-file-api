#region Code
using DevExpress.AIIntegration.Docs;
using DevExpress.XtraRichEdit;
using System.Globalization;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations;

namespace BlazorDemo.OfficeFileApi.Controllers.AI.DocumentOperations {
    [ApiController]
    [Route("api/ai")]
    public class WordTranslateController : WordControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public WordTranslateController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/word-translate")]
        public async Task<IActionResult> WordTranslate(IFormFile file, [FromForm] WordTranslationModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await TranslateDocument(_aiDocProcessingService, inputStream, model.TargetLanguage);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> TranslateDocument(IAIDocProcessingService service, Stream inputStream, string targetLanguage) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load the input document
            wordProcessor.LoadDocument(inputStream);

            await service.TranslateAsync(wordProcessor, new CultureInfo(targetLanguage));

            // Create output stream
            var outputStream = new MemoryStream();

            wordProcessor.SaveDocument(outputStream, DocumentFormat.Docx);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
