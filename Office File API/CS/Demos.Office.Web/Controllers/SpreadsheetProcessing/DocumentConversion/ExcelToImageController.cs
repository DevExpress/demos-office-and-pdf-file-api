#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentConversion;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.DocumentConversion {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelToImageController : SpreadsheetControllerBase {
        [HttpPost("excel-to-image/convert")]
        public async Task<IActionResult> ConvertExcelToImage(IFormFile file, [FromForm] ExcelToImageConversionModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetImageFormat = ImageFormatHelper.GetImageFileFormat(model.ImageFormat);

            // Convert worksheets to images
            var imageStreams = ConvertToImages(inputStream, targetImageFormat, model.ExportMode == ImageExportMode.ActiveSheet, model.Resolution, model.ExportGridlines);

            if(imageStreams == null || imageStreams.Count == 0)
                return StatusCode(500, new { Success = false, Message = "Failed to convert document to images" });

            if(imageStreams.Count == 1) {
                // Single image - return directly
                var fileName = "page_image";
                return CreateOutputFile((Stream)imageStreams[0], model.ImageFormat, fileName);
            } else {
                // Multiple images - create ZIP file
                var zipFileName = $"{model.OutputName}.zip";
                return CreateZipOutputFile(imageStreams, model.ImageFormat, "page", zipFileName);
            }
        }
        #region Code
        IReadOnlyList<Stream> ConvertToImages(Stream inputStream, ImageFileFormat imageFormat, bool exportActiveWorksheetOnly,
            int resolution, bool exportGridlines) {
            using var workbook = new Workbook();

            // Load the input document
            workbook.LoadDocument(inputStream);

            // Configure image export options
            var imageExportOptions = new RangeImageOptions {
                Resolution = resolution,
                ExportGridlines = exportGridlines,
                GridlineColor = Color.Gray,
            };

            // Get worksheets to export
            var worksheetsToExport = new List<Worksheet>();
            if(exportActiveWorksheetOnly) {
                var activeSheet = workbook.Worksheets.ActiveWorksheet;
                worksheetsToExport.Add(activeSheet);
            } else
                worksheetsToExport = workbook.Worksheets.ToList();

            // Export document pages to image streams
            var imageStreams = new List<Stream>();
            foreach(var worksheet in worksheetsToExport) {
                var worksheetImageStream = new MemoryStream();
                worksheet.GetUsedRange().ExportToImage(worksheetImageStream, imageFormat, imageExportOptions);
                worksheetImageStream.Position = 0;
                imageStreams.Add(worksheetImageStream);
            }
            return imageStreams.AsReadOnly();
        }
        #endregion Code
    }
}
