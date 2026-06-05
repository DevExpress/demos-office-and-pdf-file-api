#region Code
using DevExpress.AIIntegration;
using DevExpress.AIIntegration.Docs;
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
using DevExpress.Office.Utils;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation;
using BlazorDemo.OfficeFileApi.Models.AI.DocumentOperations;

namespace BlazorDemo.OfficeFileApi.Controllers.AI.DocumentOperations {
    [ApiController]
    [Route("api/ai")]
    public class PresentationSummarizeController : PresentationControllerBase {
        private readonly IAIDocProcessingService _aiDocProcessingService;
        public PresentationSummarizeController(IAIDocProcessingService aiDocProcessingService) {
            _aiDocProcessingService = aiDocProcessingService;
        }
        [HttpPost("document-operations/presentation-summarize")]
        public async Task<IActionResult> PresentationSummarize(IFormFile file, [FromForm] PresentationSummarizationModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var outputStream = await SummarizeDocument(_aiDocProcessingService, inputStream, model.SummarizationMode, model.OutputMode);

            return CreateOutputFile(outputStream, model.OutputMode == OutputMode.StandaloneFile ? "txt" : "pptx", model.OutputName);
        }

        #region Code
        async Task<Stream> SummarizeDocument(IAIDocProcessingService service, Stream inputStream, SummarizationMode summarizationMode, OutputMode outputMode) {
            using var presentation = new Presentation(inputStream);

            var result = await service.SummarizeAsync(presentation, summarizationMode);

            // Create output stream
            var outputStream = new MemoryStream();

            switch(outputMode) {
                case OutputMode.InsertAboveContent:
                    InsertToBeginning(presentation, result);
                    presentation.SaveDocument(outputStream);
                    break;
                case OutputMode.InsertBelowContent:
                    InsertToEnd(presentation, result);
                    presentation.SaveDocument(outputStream);
                    break;
                case OutputMode.StandaloneFile:
                    CreateStandaloneFile(outputStream, result);
                    break;
                default:
                    throw new InvalidOperationException("Invalid output mode");
            }

            outputStream.Position = 0;
            return outputStream;
        }

        void InsertToBeginning(Presentation presentation, string text) {
            // Insert a new slide at the beginning
            var slide = CreateSummarySlide(presentation, text);
            presentation.Slides.Insert(0, slide);
        }

        void InsertToEnd(Presentation presentation, string text) {
            // Insert a new slide at the end
            var slide = CreateSummarySlide(presentation, text);
            presentation.Slides.Add(slide);
        }

        Slide CreateSummarySlide(Presentation presentation, string text) {
            SlideMaster slideMaster = presentation.SlideMasters[0];
            SlideLayout slideLayout = slideMaster.Layouts.GetOrCreate(SlideLayoutType.Object);
            Slide slide = new Slide(slideLayout);
            foreach(Shape shape in slide.Shapes) {
                if(shape.PlaceholderSettings.Type is PlaceholderType.Title)
                    shape.TextArea = new TextArea("Document Summary");
                if(shape.PlaceholderSettings.Type is PlaceholderType.Body) {
                    shape.TextArea = new TextArea(text);
                    foreach(var paragraph in shape.TextArea.Paragraphs)
                        paragraph.Properties.ListBullet = ListBullet.None;
                }
            }
            return slide;
        }

        void CreateStandaloneFile(Stream outputStream, string text) {
            var writer = new StreamWriter(outputStream);
            writer.Write(text);
            writer.Flush();
        }
        #endregion Code
    }
}
