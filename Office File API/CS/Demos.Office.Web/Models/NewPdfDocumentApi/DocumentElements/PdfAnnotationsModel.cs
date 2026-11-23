using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.PdfProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewPdfDocumentApi.DocumentElements {
    public class PdfAnnotationsModel : PdfModelBase, IValidatableObject {
        public NewPdfAnnotationActionType ActionType { get; set; } = NewPdfAnnotationActionType.Add;
        internal readonly List<FormatInfo> ActionTypes = new() {
            new() { Text = "Add",     Value = NewPdfAnnotationActionType.Add },
            new() { Text = "Flatten", Value = NewPdfAnnotationActionType.Flatten },
            new() { Text = "Clear",   Value = NewPdfAnnotationActionType.Clear }
        };
        public NewPdfAnnotationOperationType OperationType { get; set; } = NewPdfAnnotationOperationType.Link;
        internal readonly List<FormatInfo> OperationTypes = new() {
            new() { Text = "Link",        Value = NewPdfAnnotationOperationType.Link },
            new() { Text = "Free Text",   Value = NewPdfAnnotationOperationType.FreeText },
            new() { Text = "Stamp",       Value = NewPdfAnnotationOperationType.Stamp },
            new() { Text = "Sticky Note", Value = NewPdfAnnotationOperationType.StickyNote }
        };
        public string Text { get; set; } = "Sample Text";
        public string SearchText { get; set; } = "DevExpress";
        public string AnnotationColor { get; set; }
        internal object Color { get; set; } = "#CC5B00";
        public string StampIconName { get; set; } = "Approved";
        internal readonly List<string> StampIconNames = new() { "Approved", "Draft", "Confidential", "Final" };
        public string StickyNoteIconName { get; set; } = "Note";
        internal readonly List<string> StickyNoteIconNames = new() { "Note", "Comment", "Cross", "Star" };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if(ActionType == NewPdfAnnotationActionType.Add && OperationType == NewPdfAnnotationOperationType.Link && string.IsNullOrWhiteSpace(SearchText)) {
                yield return new ValidationResult(
                        "A search term is required to create a link associated with text",
                        [nameof(SearchText)]);
                yield break;
            }
        }
    }

    public enum NewPdfAnnotationActionType { Add, Flatten, Clear }

    public enum NewPdfAnnotationOperationType { Link, FreeText, Stamp, StickyNote }
}
