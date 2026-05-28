using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.SearchExtract {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfExtractTextController : PdfControllerBase {
        [HttpPost("pdf-search-extract/extract-text")]
        public async Task<IActionResult> GetPdfText(IFormFile file, [FromForm] PdfExtractTextModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange).Select(x => x + 1);

            var outputStream = GetText(documentStream, pageRange);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GetText(Stream documentStream, IEnumerable<int> pageRange) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);
            var outputStream = new MemoryStream();

            string text = string.Empty;

            if(!pageRange.Any())
                text = processor.GetText();
            else
                foreach(var index in pageRange)
                    if(index < 1 || index > processor.Document.Pages.Count)
                        continue;
                    else
                        text += processor.GetPageText(index);

            var writer = new StreamWriter(outputStream);
            writer.Write(text);
            writer.Flush();

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
