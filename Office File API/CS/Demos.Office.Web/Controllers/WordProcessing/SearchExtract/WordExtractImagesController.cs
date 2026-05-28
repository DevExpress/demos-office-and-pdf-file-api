using BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.Drawing;
using DevExpress.Office.Services;
using DevExpress.Spreadsheet.Charts;
using DevExpress.XtraRichEdit;
using DevExpress.XtraSpreadsheet.Services;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.SearchExtract {
    [ApiController]
    [Route("api/wordprocessing")]
    public class WordExtractImagesController : WordControllerBase {
        [HttpPost("extract-images/process")]
        public async Task<IActionResult> ExtractDocumentImages(IFormFile file, [FromForm] WordExtractImagesModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var imageStreams = ExtractImages(inputStream, model.ExportCharts, model.ImageFormat);

            if(imageStreams == null || imageStreams.Count == 0) {
                return StatusCode(500, new { Success = false, Message = "Failed to extract document images" });
            }
            var zipFileName = $"{model.OutputName}.zip";
            return CreateZipOutputFile(imageStreams, model.ImageFormat.ToString(), "image", zipFileName);
        }
        #region Code
        IReadOnlyList<Stream> ExtractImages(Stream inputStream, bool exportCharts, DXImageFormat imageFormat) {
            OfficeCharts.Instance.ActivateCrossPlatformCharts();
            using var wordProcessor = new RichEditDocumentServer();
            try {
                IList<Stream> imageStreams = new List<Stream>();
                wordProcessor.LoadDocument(inputStream);

                var document = wordProcessor.Document;
                foreach(var shape in document.Shapes) {
                    if(shape.Type == DevExpress.XtraRichEdit.API.Native.ShapeType.Picture) {
                        Stream imgStream = new MemoryStream();
                        shape.PictureFormat.Picture.DXImage.Save(imgStream, imageFormat);
                        imageStreams.Add(imgStream);
                    }
                    else if(exportCharts && shape.Type == DevExpress.XtraRichEdit.API.Native.ShapeType.Chart) {
                        Stream imgStream = new MemoryStream();
                        ((ChartObject)shape.ChartFormat.Chart).GetImage().DXImage.Save(imgStream, imageFormat);
                        imageStreams.Add(imgStream);
                    }
                }
                return imageStreams.AsReadOnly();
            }
            catch {
                return Array.Empty<Stream>();
            }
        }
        #endregion Code
    }
}
