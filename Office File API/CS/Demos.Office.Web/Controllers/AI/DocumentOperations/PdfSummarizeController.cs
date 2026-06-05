#region Code
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Pdf;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
using BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations;

namespace BlazorDemo.OfficeFileApi.Controllers.AI.DocumentOperations {
    [ApiController]
    [Route("api/ai")]
    public class PdfSummarizeController : PdfControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public PdfSummarizeController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/pdf-summarize")]
        public async Task<IActionResult> PdfSummarize(IFormFile file, [FromForm] PdfSummarizationModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await SummarizeDocument(_aiDocProcessingService, inputStream, model.SummarizationMode);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        async Task<Stream> SummarizeDocument(IAIDocProcessingService service, Stream inputStream, SummarizationMode summarizationMode) {
            using var processor = new PdfDocumentProcessor();

            // Load the input document
            processor.LoadDocument(inputStream);

            var result = await service.SummarizeAsync(processor, summarizationMode);

            // Create output stream
            var outputStream = new MemoryStream();

            var writer = new StreamWriter(outputStream);
            writer.Write(result);
            writer.Flush();

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
