using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.OCR.Recognition;
using BlazorDemo.OfficeFileApi.Services;
#region Code
using System.Text;
using DevExpress.Docs.Ocr;
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.OCR.Recognition {
    [ApiController]
    [Route("api/ocr")]
    public class OcrRecognizeRegionController : OcrControllerBase {
        readonly OcrProviderService ocrProviders;

        public OcrRecognizeRegionController(OcrProviderService ocrProviders) {
            this.ocrProviders = ocrProviders;
        }

        [HttpPost("recognition/recognize-region")]
        public async Task<IActionResult> RecognizeRegion(IFormFile file, [FromForm] OcrRecognizeRegionModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = RecognizeRegion(documentStream, model.Region, model.OutputFormat, HttpContext.RequestAborted);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        static readonly Dictionary<OcrPageRegion, System.Drawing.RectangleF> PredefinedRegions = new() {
            [OcrPageRegion.Header] = new(0, 762, 595, 79),
            [OcrPageRegion.Body] = new(0, 140, 595, 610),
            [OcrPageRegion.Footer] = new(0, 0, 595, 128)
        };

        MemoryStream RecognizeRegion(Stream documentStream, OcrPageRegion region, string outputFormat, CancellationToken cancellationToken) {
            var provider = ocrProviders.GetProvider(detectPageOrientation: false);
            var engine = new OcrEngine(provider);

            using var document = new PdfDocument(documentStream);

            var options = new PdfDocumentRecognitionOptions {
                Regions = new[] { PredefinedRegions[region] }
            };

            var result = engine.RecognizePdfDocument(document, options, cancellationToken: cancellationToken);

            var outputStream = new MemoryStream();
            if(outputFormat == "pdf") {
                using var pdf = result.ToPdf();
                pdf.Save(outputStream);
            }
            else {
                var pages = result.PageResults.ToList();
                var content = outputFormat == "json"
                    ? Serialize(new { region = region.ToString(), pages = ToPages(pages) })
                    : ToText(pages);
                outputStream.Write(new UTF8Encoding(false).GetBytes(content));
            }
            return outputStream;
        }
        #endregion
    }
}
