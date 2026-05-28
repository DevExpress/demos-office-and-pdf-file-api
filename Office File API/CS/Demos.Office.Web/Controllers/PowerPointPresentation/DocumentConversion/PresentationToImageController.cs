using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentConversion;
#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Drawing;
#endregion Code
using Microsoft.AspNetCore.Mvc;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentConversion {
    [ApiController]
    [Route("api/presentation")]
    public class PresentationToImageController : PresentationControllerBase {
        [HttpPost("presentation-to-image/convert")]
        public async Task<IActionResult> ConvertPresentationToImage(IFormFile file, [FromForm] PresentationToImageConversionModel model) {
            if(file == null || file.Length == 0)
                return BadRequest(new { Success = false, Message = "No file uploaded" });

            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var slideIndexes = RangeHelper.GetIndexesByRange(model.ExportMode, model.SlideRange);
            var imageStreams = ConvertToImages(inputStream, model.ImageFormat, model.Resolution, slideIndexes);

            if(imageStreams == null || imageStreams.Count == 0) {
                return StatusCode(500, new { Success = false, Message = "Failed to convert document to images" });
            }

            if(imageStreams.Count == 1) {
                // Single image - return directly
                var fileName = "slide_image";
                return CreateOutputFile((Stream)imageStreams[0], model.ImageFormat.ToString(), fileName);
            } else {
                // Multiple images - create ZIP file
                var zipFileName = $"{model.OutputName}.zip";
                return CreateZipOutputFile(imageStreams, model.ImageFormat.ToString(), "slide", zipFileName);
            }
        }
        #region Code
        IReadOnlyList<Stream> ConvertToImages(Stream inputStream, DXImageFormat imageFormat, int resolution, List<int> slideIndexes) {
            var outputStreams = new List<Stream>();
            using var presentation = new Presentation(inputStream);
            var images = presentation.ExportToImages(new ImageExportOptions { Resolution = resolution }, slideIndexes.ToArray());
            foreach(var image in images) {
                var outputStream = new MemoryStream();
                image.Save(outputStream, imageFormat);
                outputStream.Position = 0;
                outputStreams.Add(outputStream);
            }
            return outputStreams.AsReadOnly();
        }
        #endregion Code
    }
}
