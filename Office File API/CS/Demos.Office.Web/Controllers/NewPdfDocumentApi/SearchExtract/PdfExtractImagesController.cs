using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.SearchExtract;
using BlazorDemo.OfficeFileApi.Services;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
using DevExpress.Drawing;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.SearchExtract {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
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
            using var doc = new PdfDocument(documentStream);
            var result = new List<Stream>();

            if(!pageRange.Any())
                pageRange = Enumerable.Range(0, doc.Pages.Count);

            foreach(var index in pageRange) {
                if(index < 0 || index >= doc.Pages.Count)
                    continue;

                foreach(var fragment in doc.Pages[index].Fragments) {
                    if(fragment is ImageFragment imageFragment) {
                        var imageStream = new MemoryStream();
                        imageFragment.Image.Save(imageStream, imageFormat);
                        result.Add(imageStream);
                    }
                }
            }

            return result;
        }
        #endregion
    }
}
