using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentConversion;
using BlazorDemo.OfficeFileApi.Services;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentConversion {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfExportToImageController : PdfControllerBase {
        [HttpPost("pdf-document-conversion/export-to-image")]
        public async Task<IActionResult> ExportToImage(IFormFile file, [FromForm] PdfExportToImageModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var pageRange = RangeHelper.GetIndexesByRange(model.ExportMode, model.PageRange);
            var imageStreams = ExportToImages(documentStream, model.ImageFormat, model.Resolution, pageRange);

            if(imageStreams?.Any() == false)
                return StatusCode(500, new { Success = false, Message = "Failed to export document pages to images" });

            if(imageStreams?.Count == 1)
                return CreateOutputFile(imageStreams[0], model.ImageFormat, "page_image");

            var zipFileName = $"{model.OutputName}.zip";
            return CreateZipOutputFile(imageStreams, model.ImageFormat, "page", zipFileName);
        }

        #region Code
        IReadOnlyList<Stream> ExportToImages(Stream documentStream, string imageFormat, int resolution, IEnumerable<int> pageRange) {
            using var document = new PdfDocument(documentStream);

            var pageIndexes = pageRange
                .Where(x => x >= 0 && x < document.Pages.Count)
                .Distinct()
                .ToArray();

            if(!pageIndexes.Any())
                pageIndexes = Enumerable.Range(0, document.Pages.Count).ToArray();

            var exportOptions = new ImageExportOptions() {
                Resolution = resolution,
                ImageExportType = imageFormat == "svg" ? ImageExportType.Svg : ImageExportType.Bitmap
            };
            
            var images = document.ExportToImages(exportOptions, pageIndexes);
            var targetFormat = ImageFormatHelper.GetDXImageFormat(imageFormat);

            var result = new List<Stream>();
            foreach(var image in images) {
                var imageStream = new MemoryStream();
                image.Save(imageStream, targetFormat);
                imageStream.Position = 0;
                result.Add(imageStream);
            }
            return result;
        }
        #endregion
    }
}
