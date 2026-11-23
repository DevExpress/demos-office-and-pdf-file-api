using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.OCR.Recognition;
using BlazorDemo.OfficeFileApi.Services;
#region Code
using System.Text;
using DevExpress.Docs.Ocr;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.OCR.Recognition {
    [ApiController]
    [Route("api/ocr")]
    public class OcrRecognizeImagesController : OcrControllerBase {
        readonly OcrProviderService ocrProviders;

        public OcrRecognizeImagesController(OcrProviderService ocrProviders) {
            this.ocrProviders = ocrProviders;
        }

        [HttpPost("recognition/recognize-images")]
        public async Task<IActionResult> RecognizeImages(IFormFile file, [FromForm] OcrRecognizeImagesModel model) {
            using var imageStream = new MemoryStream();
            await file.CopyToAsync(imageStream);
            imageStream.Position = 0;

            var outputStream = RecognizeImage(imageStream, model.DetectPageOrientation, model.OutputFormat, HttpContext.RequestAborted);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        MemoryStream RecognizeImage(Stream imageStream, bool detectPageOrientation, string outputFormat, CancellationToken cancellationToken) {
            var provider = ocrProviders.GetProvider(detectPageOrientation);
            var engine = new OcrEngine(provider);

            using var image = DXImage.FromStream(imageStream);

            var results = engine.RecognizeImages(new[] { image }, cancellationToken: cancellationToken);

            var outputStream = new MemoryStream();
            if(outputFormat == "pdf") {
                using var document = results.ToPdf(new PdfGenerationOptions { RotatePages = detectPageOrientation });
                document.Save(outputStream);
            }
            else {
                var fragments = results.ImageResults.Single().TextFragments;
                var content = outputFormat == "json"
                    ? Serialize(new { fragments = ToFragments(fragments) })
                    : ToText(fragments);
                outputStream.Write(new UTF8Encoding(false).GetBytes(content));
            }
            return outputStream;
        }
        #endregion
    }
}
