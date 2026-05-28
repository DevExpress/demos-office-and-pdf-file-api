using System.Drawing;
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.ManagePages;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.ManagePages {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
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
            using var doc = new PdfDocument(documentStream);

            if(!pageRange.Any())
                pageRange = Enumerable.Range(0, doc.Pages.Count);

            foreach(var index in pageRange) {
                if(index < 0 || index >= doc.Pages.Count)
                    continue;
                var clearRegion = new OrientedRectangle(doc.Pages[index].MediaBox);
                doc.ClearContent(index, [clearRegion], clearText, clearImages, clearGraphics, clearAnnotations);
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
