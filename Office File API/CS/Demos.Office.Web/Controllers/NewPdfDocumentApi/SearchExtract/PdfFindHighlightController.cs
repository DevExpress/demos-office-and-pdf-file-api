using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.SearchExtract;
using BlazorDemo.OfficeFileApi.Controllers.PdfProcessing;
#region Code
using DevExpress.Docs.Pdf;
#endregion

namespace BlazorDemo.OfficeFileApi.Controllers.NewPdfDocumentApi.SearchExtract {
    [ApiController]
    [Route("api/newpdfdocumentapi")]
    public class PdfFindHighlightController : PdfControllerBase {
        [HttpPost("pdf-search-extract/find-highlight-text")]
        public async Task<IActionResult> FindHighlightPdfText(IFormFile file, [FromForm] PdfFindHighlightModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var color = System.Drawing.ColorTranslator.FromHtml(model.AnnotationColor);
            var outputStream = FindHighlightText(documentStream, model.SearchText, color, model.AnnotationType);
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream FindHighlightText(Stream documentStream, string searchText, System.Drawing.Color color, TextMarkupAnnotationType annotationType) {
            using var doc = new PdfDocument(documentStream);

            foreach(var result in doc.FindText(searchText)) {
                foreach(var match in result.Matches)  {
                    var annotation = new TextMarkupAnnotation(match.MatchFragments.Select(f => f.Rectangle).ToArray()) {
                            Color  = PdfColor.FromRgb(color.R, color.G, color.B),
                            CreationDate = DateTime.UtcNow,
                            Title  = "DX Demo",
                            Type   = annotationType
                        };
                    doc.Pages[result.PageIndex].Annotations.Add(annotation);
                }
            }
            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
