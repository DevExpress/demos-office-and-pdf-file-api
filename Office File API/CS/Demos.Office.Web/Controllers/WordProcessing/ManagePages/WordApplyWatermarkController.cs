using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ManagePages {
    [ApiController]
    [Route("api/wordprocessing/watermark")]
    public class WordApplyWatermarkController : WordControllerBase {
        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentWatermarkTemplate, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "Watermark_Template.docx");
        }
        [HttpPost("process")]
        public async Task<IActionResult> ProcessWatermarkTemplate([FromForm] WordWatermarkModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentWatermarkTemplate, FileMode.Open, FileAccess.Read);
            Stream inputStream = new MemoryStream();
            templateStream.CopyTo(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Process watermark on the template
            Stream outputStream = new MemoryStream();
            if (model.WatermarkType == WatermarkType.Text)
                outputStream = ApplyTextWatermark(inputStream, model.WatermarkText, model.FontName, model.FontSize, ColorTranslator.FromHtml(model.TextColor.ToString()),
                    model.Layout, model.Semitransparent, targetFormat);
            else
 {
                using var imageFileStream = new FileStream(DemoDocumentPaths.WordFileAttachmentWatermarkImageFile, FileMode.Open);
                Stream imageStream = new MemoryStream();
                imageFileStream.CopyTo(imageStream);
                imageStream.Position = 0;
                outputStream = ApplyImageWatermark(inputStream, imageStream, model.ScaleFactor, model.Washout, targetFormat);
            }

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ApplyTextWatermark(Stream inputStream, string watermarkText, string fontName, int fontSize, Color textColor,
            WatermarkLayout layout, bool semitransparent, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(inputStream);
            Document document = wordProcessor.Document;

            TextWatermarkOptions options = new() {
                FontFamily = fontName,
                FontSize = fontSize,
                Color = textColor,
                Layout = layout,
                Semitransparent = semitransparent
            };
            document.WatermarkManager.SetText(watermarkText, options);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream ApplyImageWatermark(Stream inputStream, Stream imageStream, double imageScale, bool washout, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            wordProcessor.LoadDocument(inputStream);
            Document document = wordProcessor.Document;

            ImageWatermarkOptions options = new() {
                Scale = imageScale,
                Washout = washout
            };
            DocumentImageSource imageSource = DocumentImageSource.FromStream(imageStream);
            document.WatermarkManager.SetImage(imageSource, options);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #endregion Code
    }
}
