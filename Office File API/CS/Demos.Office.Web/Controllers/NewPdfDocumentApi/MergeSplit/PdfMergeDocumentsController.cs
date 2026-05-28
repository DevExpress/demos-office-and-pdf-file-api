using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.MergeSplit;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.MergeSplit {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfMergeDocumentsController : PdfControllerBase {
        [HttpPost("pdf-merge-split/merge")]
        public async Task<IActionResult> MergePdfDocuments(IList<IFormFile> files, [FromForm] PdfMergeDocumentsModel model) {
            var documentStreams = await Task.WhenAll(files.Select(async x => {
                var stream = new MemoryStream();
                await x.CopyToAsync(stream);
                stream.Position = 0;
                return stream;
            }));

            var outputStream = MergeDocuments(documentStreams);
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream MergeDocuments(IEnumerable<Stream> documentStreams) {
            using var doc = new PdfDocument();

            foreach(var stream in documentStreams)
                doc.AppendDocument(stream);

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
