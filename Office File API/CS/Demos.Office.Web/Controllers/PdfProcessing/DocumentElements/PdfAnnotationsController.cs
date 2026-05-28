using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.DocumentElements {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfAnnotationsController : PdfControllerBase {
        [HttpPost("document-elements/pdf-annotations-link")]
        public async Task<IActionResult> CreatePdfLink(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = CreateLink(documentStream, model.SearchText, System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        IReadOnlyList<PdfRectangle> FindText(PdfDocumentProcessor processor, string searchText) {
            var search = processor.FindText(
                    searchText,
                    new PdfTextSearchParameters() {
                        CaseSensitive = false,
                        WholeWords = false
                    });

            if(search.Status == PdfTextSearchStatus.Found)
                return search.Rectangles.Select(x => x.BoundingRectangle).ToList();
            else
                throw new InvalidOperationException();
        }

        Stream CreateLink(Stream documentStream, string searchText, System.Drawing.Color annotationColor) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var page = processor.DocumentFacade.Pages[0];
            var textRectangles = FindText(processor, searchText);

            foreach (var rectangle in textRectangles) {
                var linkAnnotation = page.AddLinkAnnotation(rectangle, "https://www.devexpress.com");
                linkAnnotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
                linkAnnotation.BorderWidth = 2;
                linkAnnotation.BorderStyle = PdfBorderStyle.Dash;
            }

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;
    
            return outputStream;
        }
        #endregion
        [HttpPost("document-elements/pdf-annotations-freetext")]
        public async Task<IActionResult> CreatePdfFreeTextAnnotation(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = CreateFreeTextAnnotation(documentStream, model.Text, System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CreateFreeTextAnnotation(MemoryStream documentStream, string text, System.Drawing.Color annotationColor) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var page = processor.DocumentFacade.Pages[0];

            var freeTextAnnotation = page.AddFreeTextAnnotation(new PdfRectangle(600, 10, 700, 110), text);
            freeTextAnnotation.Author = "DX Demo";
            freeTextAnnotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
            freeTextAnnotation.CreationDate = DateTime.UtcNow;

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
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
        Stream CreateRubberStamp(MemoryStream documentStream, string iconName) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var page = processor.DocumentFacade.Pages[0];
            var stamp = page.AddRubberStampAnnotation(new PdfPoint(600, 100), iconName);
            stamp.Author = "DX Demo";
            stamp.CreationDate = DateTime.UtcNow;

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
        [HttpPost("document-elements/pdf-annotations-stickynote")]
        public async Task<IActionResult> CreatePdfStickyNote(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = CreateStickyNote(documentStream, model.StickyNoteIconName, System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor));
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CreateStickyNote(MemoryStream documentStream, string iconName, System.Drawing.Color annotationColor) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var page = processor.DocumentFacade.Pages[0];
            var stickyNote = page.AddTextAnnotation(new PdfPoint(600, 100), iconName);
            stickyNote.Author = "DX Demo";
            stickyNote.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
            stickyNote.CreationDate = DateTime.UtcNow;

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
        [HttpPost("document-elements/pdf-annotations-flatten")]
        public async Task<IActionResult> FlattenPdfAnnotations(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = FlattenAnnotations(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FlattenAnnotations(MemoryStream documentStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            processor.DocumentFacade.FlattenAnnotations();

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
        [HttpPost("document-elements/pdf-annotations-clear")]
        public async Task<IActionResult> ClearPdfAnnotations(IFormFile file, [FromForm] PdfAnnotationsModel model) {
            using var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = ClearAnnotations(documentStream);
            outputStream.Seek(0, SeekOrigin.Begin);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ClearAnnotations(MemoryStream documentStream) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            var annotations = processor.DocumentFacade.Pages.SelectMany(p => p.Annotations).ToList();
            foreach(var annotation in annotations)
                annotation.Remove();

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);
            outputStream.Position = 0;

            return outputStream;
        }
        #endregion
    }
}
