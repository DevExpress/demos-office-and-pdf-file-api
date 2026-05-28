using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract;
using BlazorDemo.OfficeFileApi.Services;
#region Code
using DevExpress.Pdf;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.SearchExtract {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfExtractImagesController : PdfControllerBase {
        [HttpPost("pdf-search-extract/extract-images")]
        public async Task<IActionResult> ExtractImagesFromPdf(IFormFile file, [FromForm] PdfExtractImagesModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);
            var targetImageFormat = ImageFormatHelper.GetDXImageFormat(model.ImageFormat);

            var imagesStream = ExtractImages(documentStream, targetImageFormat, model.Resolution, pageRange);

            if(imagesStream?.Any() == false)
                return StatusCode(500, new { Success = false, Message = "Failed to extract images from document" });

            var zipFileName = $"{model.OutputName}.zip";
            return CreateZipOutputFile(imagesStream, model.ImageFormat, "image", zipFileName);
        }

        #region Code
        IReadOnlyList<Stream> ExtractImages(Stream documentStream, DXImageFormat imageFormat, int resolution, IEnumerable<int> pageRange) {
            using var processor = new PdfDocumentProcessor();
            var result = new List<Stream>();

            processor.LoadDocument(documentStream);

            if (!pageRange.Any())
                pageRange = Enumerable.Range(0, processor.Document.Pages.Count);

            foreach(var index in pageRange) {
                if(index < 0 || index >= processor.Document.Pages.Count)
                    continue;

                var page = processor.Document.Pages[index];

                var images = processor.GetDXImages(
                    new PdfDocumentArea(page.GetPageIndex(), page.MediaBox),
                    resolution);

                foreach(var image in images) {
                    var imageStream = new MemoryStream();
                    image.Save(imageStream, imageFormat);
                    result.Add(imageStream);
                }
            }

            return result;
        }
        #endregion
    }
}
