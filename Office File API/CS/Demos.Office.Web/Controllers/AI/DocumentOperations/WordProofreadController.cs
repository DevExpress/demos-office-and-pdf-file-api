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
    public class WordProofreadController : WordControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public WordProofreadController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/word-proofread")]
        public async Task<IActionResult> WordProofread(IFormFile file, [FromForm] WordProofreadingModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await ProofreadDocumentAsync(_aiDocProcessingService, inputStream, model.SourceLanguage);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> ProofreadDocumentAsync(IAIDocProcessingService service, Stream inputStream, string sourceLanguage) {
            using var wordProcessor = new RichEditDocumentServer();

            // Load the input document
            wordProcessor.LoadDocument(inputStream);

            await service.ProofreadAsync(wordProcessor, new CultureInfo(sourceLanguage));

            // Create output stream
            var outputStream = new MemoryStream();

            wordProcessor.SaveDocument(outputStream, DocumentFormat.Docx);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
