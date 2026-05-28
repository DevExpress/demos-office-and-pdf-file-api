using System.Drawing;
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentProtection;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;

#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentProtection {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfApplyWatermarkController : PdfControllerBase {
        [HttpPost("document-protection/apply-watermark")]
        public async Task<IActionResult> ApplyWatermarkToPdf(
            IFormFile file,
            [FromForm] PdfApplyWatermarkModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = ApplyWatermark(documentStream, model.WatermarkText, model.Rotate);
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ApplyWatermark(Stream documentStream, string text, bool rotate) {
            using var doc = new PdfDocument(documentStream);

            FormTemplate form = new() { Bounds = new RectangleF(0,0,842, 595) };
            form.AddFragment(new TextFragment() {
                Text = text,
                Location = rotate ? new PointF(300, 220) : new PointF(280, 300) ,
                ScaleX = 4,
                ScaleY = 4,
                ForegroundFill = new SolidFill(PdfColor.LightSkyBlue, 0.5),
                RotationAngle = rotate ? 45 : 0
            });

            foreach(var page in doc.Pages) {
                WatermarkAnnotation watermark = new(page.MediaBox) {
                    Appearance = new AnnotationAppearances {
                        Normal = new AnnotationAppearance {
                            DefaultForm = form
                        }
                    }
                };
                page.Annotations.Add(watermark);
            }
            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
