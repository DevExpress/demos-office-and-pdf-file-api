using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentElements {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfBookmarksController : PdfControllerBase {
        [HttpPost("pdf-bookmarks/add")]
        public async Task<IActionResult> AddBookmarksToPdf(IFormFile file, [FromForm] PdfBookmarksModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = AddBookmarks(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddBookmarks(Stream documentStream) {
            using var doc = new PdfDocument(documentStream);

            for(int i = 0; i < doc.Pages.Count; i++)
                doc.Bookmarks.Add($"Page {i + 1}", doc.Pages[i], new FitBBoxDestination());

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            return outputStream;
        }

        #endregion
        [HttpPost("pdf-bookmarks/delete")]
        public async Task<IActionResult> DeleteBookmarksFromPdf(IFormFile file, [FromForm] PdfBookmarksModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = DeleteBookmarks(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream DeleteBookmarks(Stream documentStream) {
            using var doc = new PdfDocument(documentStream);

            doc.Bookmarks.Clear();

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            return outputStream;
        }

        #endregion
    }
}
