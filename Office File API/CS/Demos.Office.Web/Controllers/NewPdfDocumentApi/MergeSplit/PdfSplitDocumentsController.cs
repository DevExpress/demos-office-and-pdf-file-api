using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.MergeSplit;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.MergeSplit {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfSplitDocumentsController : PdfControllerBase {
        [HttpPost("pdf-merge-split/split")]
        public async Task<IActionResult> SplitPdfDocuments(IFormFile file, [FromForm] PdfSplitDocumentsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);

            var outputStreams = SplitDocuments(documentStream, pageRange);

            return CreateZipOutputFile(outputStreams, "pdf", "document", model.OutputName);
        }

        #region Code
        IReadOnlyList<Stream> SplitDocuments(Stream documentStream, IEnumerable<int> pageRange) {
            using var source = new PdfDocument(documentStream);

            if(!pageRange.Any())
                pageRange = Enumerable.Range(0, source.Pages.Count);

            var result = new List<Stream>();
            foreach(var index in pageRange) {
                if(index < 0 || index >= source.Pages.Count)
                    continue;

                using var pageDoc = new PdfDocument();
                pageDoc.Pages.Add(source.Pages[index].Clone());

                var outputStream = new MemoryStream();
                pageDoc.Save(outputStream);
                outputStream.Position = 0;
                result.Add(outputStream);
            }

            return result;
        }
        #endregion
    }
}
