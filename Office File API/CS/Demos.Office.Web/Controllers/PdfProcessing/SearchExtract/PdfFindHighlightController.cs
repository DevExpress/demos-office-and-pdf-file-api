using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract;
#region Code
using DevExpress.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.PdfProcessing.SearchExtract {
    [ApiController]
    [Route("api/pdfprocessing")]
    public class PdfFindHighlightController : PdfControllerBase {
        [HttpPost("pdf-search-extract/find-highlight-text")]
        public async Task<IActionResult> FindHighlightPdfText(IFormFile file, [FromForm] PdfFindHighlightModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = FindHighlightText(documentStream, model.SearchText, model.AnnotationType, System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor));
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream FindHighlightText(MemoryStream documentStream, string searchText, PdfTextMarkupAnnotationType annotationType, System.Drawing.Color annotationColor) {
            using var processor = new PdfDocumentProcessor();

            processor.LoadDocument(documentStream);

            PdfTextSearchResults search;

            do {
                search = processor.FindText(
                    searchText,
                    new PdfTextSearchParameters() {
                        CaseSensitive = false,
                        WholeWords = false
                    });
                if(search.Status == PdfTextSearchStatus.Found) {
                    var annotation = processor.DocumentFacade.Pages[search.Page.GetPageIndex()].AddTextMarkupAnnotation(search.Rectangles, annotationType);

                    annotation.Author = "DX Demo";
                    annotation.CreationDate = DateTime.UtcNow;
                    annotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
                }
            } while(search.Status == PdfTextSearchStatus.Found);

            var outputStream = new MemoryStream();
            processor.SaveDocument(outputStream);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion
    }
}
