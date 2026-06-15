using System.Drawing;
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.DocumentElements {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfAnnotationsController : PdfControllerBase {
        [HttpPost("document-elements/pdf-annotations-link")]
        public async Task<IActionResult> CreatePdfLink(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var color = System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor);
            var outputStream = CreateLink(documentStream, model.SearchText, color);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateLink(Stream documentStream, string searchText, System.Drawing.Color color) {
            using var doc = new PdfDocument(documentStream);

            foreach(var result in doc.FindText(searchText)) {
                foreach(var match in result.Matches) {
                    var bounds = match.MatchFragments.Select(f => f.Rectangle.BoundingBox)
                                      .Aggregate(RectangleF.Union);
                    var link = new LinkAnnotation(bounds) {
                        Action = new UriAction { Uri = "https://www.devexpress.com" },
                        Color = PdfColor.FromRgb(color.R, color.G, color.B),
                        BorderStyle = new AnnotationBorderStyle { Width = 2, Style = BorderStyle.Dashed }
                    };
                    doc.Pages[result.PageIndex].Annotations.Add(link);
                }
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion

        [HttpPost("document-elements/pdf-annotations-freetext")]
        public async Task<IActionResult> CreatePdfFreeTextAnnotation(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var color = System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor);
            var outputStream = CreateFreeTextAnnotation(documentStream, model.Text, color);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateFreeTextAnnotation(Stream documentStream, string text, System.Drawing.Color color) {
            using var doc = new PdfDocument(documentStream);

            var page = doc.Pages[0];
            page.Annotations.Add(new FreeTextAnnotation(new RectangleF(450, 20, 130, 100)) {
                Content = text,
                Title = "DX Demo",
                Color = PdfColor.FromRgb(color.R, color.G, color.B),
                CreationDate = DateTimeOffset.UtcNow
            });

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion

        [HttpPost("document-elements/pdf-annotations-stamp")]
        public async Task<IActionResult> CreatePdfRubberStamp(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = CreateRubberStamp(documentStream, model.StampIconName);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateRubberStamp(Stream documentStream, string iconName) {
            using var doc = new PdfDocument(documentStream);

            var page = doc.Pages[0];
            page.Annotations.Add(new RubberStampAnnotation(
                    new RectangleF(350, 80, 230, 55),
                    RubberStampAnnotationIconName.FromName(iconName)) {
                Title = "DX Demo",
                CreationDate = DateTimeOffset.UtcNow
            });

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion

        [HttpPost("document-elements/pdf-annotations-stickynote")]
        public async Task<IActionResult> CreatePdfStickyNote(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var color = System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor);
            var outputStream = CreateStickyNote(documentStream, model.StickyNoteIconName, color);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream CreateStickyNote(Stream documentStream, string iconName, System.Drawing.Color color) {
            using var doc = new PdfDocument(documentStream);

            var page = doc.Pages[0];
            page.Annotations.Add(new TextAnnotation(new RectangleF(450, 150, 20, 20)) {
                IconName = Enum.Parse<TextAnnotationIconName>(iconName),
                Color = PdfColor.FromRgb(color.R, color.G, color.B),
                Content = "This is a Sticky Note",
                Title = "DX Demo",
                CreationDate = DateTimeOffset.UtcNow
            });

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion

        [HttpPost("document-elements/pdf-annotations-clear")]
        public async Task<IActionResult> ClearAnnotations(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = ClearAllAnnotations(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream ClearAllAnnotations(Stream documentStream) {
            using var doc = new PdfDocument(documentStream);

            foreach(var page in doc.Pages)
                page.Annotations.Clear();

            var outputStream = new MemoryStream();
            doc.Save(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
