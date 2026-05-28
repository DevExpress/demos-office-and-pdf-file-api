using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing;
using DevExpress.Docs.Pdf;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.SearchExtract {
    public class PdfFindHighlightModel : PdfModelBase, IValidatableObject {
        [Required(ErrorMessage = "Add a search term for annotation")]
        [MaxLength(15, ErrorMessage = "Text must not exceed 15 symbols")]
        public string SearchText { get; set; } = "DevExpress";
        public TextMarkupAnnotationType AnnotationType { get; set; } = TextMarkupAnnotationType.Highlight;
        internal readonly List<FormatInfo> AnnotationTypes = new() {
            new() { Text = "Highlight",   Value = TextMarkupAnnotationType.Highlight },
            new() { Text = "Squiggly",    Value = TextMarkupAnnotationType.Squiggly },
            new() { Text = "Strike Out",  Value = TextMarkupAnnotationType.StrikeOut },
            new() { Text = "Underline",   Value = TextMarkupAnnotationType.Underline }
        };
        public string AnnotationColor { get; set; }
        internal object Color { get; set; } = "#FF0000";
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();
            if(string.IsNullOrWhiteSpace(Color?.ToString()))
                results.Add(new ValidationResult("Select a color for annotation", [nameof(AnnotationColor)]));
            return results;
        }
    }
}
