using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using DevExpress.Pdf;
using DevExpress.XtraSpreadsheet.PrintLayoutEngine;

namespace BlazorDemo.OfficeFileApi.Models.PdfProcessing.DocumentElements {
    public class PdfAnnotationsModel : PdfModelBase, IValidatableObject {
        public PdfAnnotationOperationType OperationType { get; set; } = PdfAnnotationOperationType.Link;
        internal readonly List<FormatInfo> OperationsTypes = new() {
            new() { Text = "Link", Value = PdfAnnotationOperationType.Link },
            new() { Text = "Free Text", Value = PdfAnnotationOperationType.FreeText },
            new() { Text = "Stamp", Value = PdfAnnotationOperationType.Stamp },
            new() { Text = "Sticky Note", Value = PdfAnnotationOperationType.StickyNote }
        };
        public string Text { get; set; } = "Sample Text";
        public string SearchText { get; set; } = "DevExpress";
        public string AnnotationColor { get; set; }
        internal object Color { get; set; } = "#CC5B00";
        public string StampIconName { get; set; } = PdfRubberStampAnnotationIconName.Approved;
        internal readonly List<string> StampIconNames = new() {
            PdfRubberStampAnnotationIconName.Approved,
            PdfRubberStampAnnotationIconName.Draft,
            PdfRubberStampAnnotationIconName.Confidential,
            PdfRubberStampAnnotationIconName.Final
        };
        public string StickyNoteIconName { get; set; } = PdfTextAnnotationIconName.Note;
        internal readonly List<string> StickyNoteIconNames = new() {
            PdfTextAnnotationIconName.Note,
            PdfTextAnnotationIconName.Comment,
            PdfTextAnnotationIconName.Cross,
            PdfTextAnnotationIconName.Star
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if (OperationType == PdfAnnotationOperationType.Link && string.IsNullOrWhiteSpace(SearchText)) {
                yield return new ValidationResult(
                        "A search term is required to create a link associated with text",
                        [nameof(SearchText)]);
                yield break;
            }
        }
    }
    public enum PdfAnnotationOperationType {
        Link, FreeText, Stamp, StickyNote
    }
}
