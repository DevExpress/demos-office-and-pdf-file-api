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
    public class OcrRecognizePdfController : OcrControllerBase {
        readonly OcrProviderService ocrProviders;

        public OcrRecognizePdfController(OcrProviderService ocrProviders) {
            this.ocrProviders = ocrProviders;
        }

        [HttpPost("recognition/recognize-pdf")]
        public async Task<IActionResult> RecognizePdf(IFormFile file, [FromForm] OcrRecognizePdfModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);
            var outputStream = RecognizePages(documentStream, model.RecognitionMode, model.DetectPageOrientation,
                pageRange, model.OutputFormat, HttpContext.RequestAborted);
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        MemoryStream RecognizePages(Stream documentStream, PdfDocumentRecognitionMode recognitionMode,
            bool detectPageOrientation, IEnumerable<int> pageRange, string outputFormat, CancellationToken cancellationToken) {
            var provider = ocrProviders.GetProvider(detectPageOrientation);
            var engine = new OcrEngine(provider);

            using var document = new PdfDocument(documentStream);

            var pageIndexes = pageRange.Where(index => index >= 0 && index < document.Pages.Count).Distinct().ToArray();

            var options = new PdfDocumentRecognitionOptions { RecognitionMode = recognitionMode };

            var result = engine.RecognizePdfDocument(document, options, cancellationToken: cancellationToken, pageIndices: pageIndexes);

            var outputStream = new MemoryStream();
            if(outputFormat == "pdf") {
                using var pdf = result.ToPdf(new PdfGenerationOptions { RotatePages = detectPageOrientation });
                pdf.Save(outputStream);
            }
            else {
                var pages = result.PageResults.ToList();
                var content = outputFormat == "json"
                    ? Serialize(new { pages = ToPages(pages) })
                    : ToText(pages);
                outputStream.Write(new UTF8Encoding(false).GetBytes(content));
            }
            return outputStream;
        }
        #endregion
    }
}
