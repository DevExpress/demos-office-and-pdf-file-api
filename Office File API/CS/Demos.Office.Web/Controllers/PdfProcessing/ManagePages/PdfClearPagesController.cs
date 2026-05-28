using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.ManagePages {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfClearPagesController : PdfControllerBase {
        [HttpPost("manage-pages/clear")]
        public async Task<IActionResult> ClearPdfPages(IFormFile file, [FromForm] PdfClearPagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);

            var outputStream = ClearPages(
                documentStream,
                model.ClearAnnotations,
                model.ClearGraphics,
                model.ClearImages,
                model.ClearText,
                pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ClearPages(Stream documentStream, bool clearAnnotations, bool clearGraphics,
            bool clearImages, bool clearText, IEnumerable<int> pageRange) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            if(!pageRange.Any())
                pageRange = Enumerable.Range(0, processor.Document.Pages.Count);

            foreach(var index in pageRange) {
                if(index < 0 || index >= processor.Document.Pages.Count)
                    continue;

                var page = processor.Document.Pages[index];

                var options = new PdfClearContentOptions() {
                    ClearAnnotations = clearAnnotations,
                    ClearGraphics = clearGraphics,
                    ClearImages = clearImages,
                    ClearText = clearText
                };

                processor.DocumentFacade.Pages[index].ClearContent(page.MediaBox, options);
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion
    }
}
