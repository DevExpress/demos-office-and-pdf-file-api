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
    public class PdfRedactionController : PdfControllerBase {
        [HttpPost("document-protection/add-redaction")]
        public async Task<IActionResult> AddRedactionToPdf(IFormFile file, [FromForm] PdfRedactionModel model) {
            var documentStream = new MemoryStream();
            await file.CopyToAsync(documentStream);
            documentStream.Position = 0;

            var outputStream = AddRedaction(documentStream, model.SearchText, model.RedactionText, model.ApplyRedaction);
            outputStream.Position = 0;

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddRedaction(Stream documentStream, string searchText, string redactionText, bool applyRedaction) {
            using var doc = new PdfDocument(documentStream);

            var searchResults = doc.FindText(searchText, new TextSearchOptions());

            foreach(var result in searchResults) {
                List<RedactionAnnotation> annotations = new List<RedactionAnnotation>();
                foreach(var match in result.Matches) {
                    var rects = match.MatchFragments
                        .Select(f => f.Rectangle).ToArray();
                    var annotation = new RedactionAnnotation(RectangleF.Empty) {
                        Geometry = new RedactionGeometry(rects),
                        FillColor  = PdfColor.Black,
                        Color  = PdfColor.Red,
                        OverlayText = string.IsNullOrEmpty(redactionText) ? null : redactionText,
                        CreationDate = DateTime.UtcNow,
                        TextJustification = TextJustification.LeftJustified,
                        RepeatText = true,
                        TextAppearance = new TextAppearance() {
                            FontSize = 0,
                            Fill = SolidFill.White,
                        }};
                    annotations.Add(annotation);
                }
                if(applyRedaction)
                    doc.ApplyRedaction(result.PageIndex, annotations.ToArray());
                else
                    foreach(var ann in annotations)
                        doc.Pages[result.PageIndex].Annotations.Add(ann);
            }

            var outputStream = new MemoryStream();
            doc.Save(outputStream);

            return outputStream;
        }
        #endregion
    }
}
