using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.DocumentElements {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfBookmarksController : PdfControllerBase {
        [HttpPost("document-elements/pdf-bookmarks-add")]
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
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            foreach (var page in processor.DocumentFacade.Pages) {
                var destination = page.CreateFitBBoxDestination();
                var bookmark = new PdfBookmark() { Destination = destination, Title = $"Page {destination.PageIndex + 1}" };
                processor.Document.Bookmarks.Add(bookmark);
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
        [HttpPost("document-elements/pdf-bookmarks-delete")]
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
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            processor.Document.Bookmarks.Clear();

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
