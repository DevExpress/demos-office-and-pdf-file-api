using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.ManagePages;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.ManagePages {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfRotateResizePagesController : PdfControllerBase {
        [HttpPost("manage-pages/rotate-resize")]
        public async Task<IActionResult> RotateResizePdfPage(
            IFormFile file,
            [FromForm] PdfRotateResizePagesModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);

            var outputStream = RotateResizePages(
                documentStream,
                pageRange,
                model.Rotate,
                new PdfRectangle(0, 0, model.Width, model.Height));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream RotateResizePages(Stream documentStream, IEnumerable<int> pageRange, int rotate, PdfRectangle pageSize) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            if (!pageRange.Any())
            pageRange = Enumerable.Range(0, processor.Document.Pages.Count);
            foreach (var index in pageRange) {
                if (index < 0 || index >= processor.Document.Pages.Count)
                    continue;

                var page = processor.Document.Pages[index];

                page.Resize(pageSize, PdfContentHorizontalAlignment.Center, PdfContentVerticalAlignment.Center);
                page.Rotate = rotate;
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            return outputStream;
        }
        #endregion
    }
}
