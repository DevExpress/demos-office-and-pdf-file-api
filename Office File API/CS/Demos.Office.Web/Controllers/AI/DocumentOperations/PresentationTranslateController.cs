#region Code
using DevExpress.AIIntegration.Docs;
using DevExpress.Docs.Presentation;
using System.Globalization;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation;
using BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations;

namespace BlazorDemo.OfficeFileApi.Controllers.AI.DocumentOperations {
    [ApiController]
    [Route("api/ai")]
    public class PresentationTranslateController : PresentationControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public PresentationTranslateController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/presentation-translate")]
        public async Task<IActionResult> PresentationTranslate(IFormFile file, [FromForm] PresentationTranslationModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await TranslateDocument(_aiDocProcessingService, inputStream, model.TargetLanguage);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> TranslateDocument(IAIDocProcessingService service, Stream inputStream, string targetLanguage) {
            using var presentation = new Presentation(inputStream);

            await service.TranslateAsync(presentation, new CultureInfo(targetLanguage));

            // Create output stream
            var outputStream = new MemoryStream();

            presentation.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
