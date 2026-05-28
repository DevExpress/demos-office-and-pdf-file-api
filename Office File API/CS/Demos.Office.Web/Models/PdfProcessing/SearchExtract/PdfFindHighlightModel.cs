using System.ComponentModel.DataAnnotations;
using DevExpress.Pdf;
using DevExpress.XtraRichEdit.Model;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.SearchExtract {
    public class PdfFindHighlightModel : PdfModelBase, IValidatableObject {
        [Required(ErrorMessage = "Add a search term for annotation")]
        [MaxLength(15, ErrorMessage = "Text must not exceed 15 symbols")]
        public string SearchText { get; set; } = "DevExpress";
        public PdfTextMarkupAnnotationType AnnotationType { get; set; } = PdfTextMarkupAnnotationType.Highlight;
        internal List<object> AnnotationTypes { get; } = new() {
            new { Text = "Highlight", Value =  PdfTextMarkupAnnotationType.Highlight},
            new { Text = "Squiggly Line", Value =  PdfTextMarkupAnnotationType.Squiggly},
            new { Text = "Strike Through", Value =  PdfTextMarkupAnnotationType.StrikeOut},
            new { Text = "Underline", Value =  PdfTextMarkupAnnotationType.Underline}
        };
        public string AnnotationColor { get; set; }
        internal object Color { get; set; } = "#CC5B00";
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            if(string.IsNullOrWhiteSpace(Color?.ToString()))
                results.Add(new ValidationResult("Select a color for annotation", [nameof(AnnotationColor)]));

            return results;
        }
    }
}
