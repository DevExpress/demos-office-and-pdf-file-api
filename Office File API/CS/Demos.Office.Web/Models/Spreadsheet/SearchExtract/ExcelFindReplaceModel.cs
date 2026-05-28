using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.SearchExtract {
    public class ExcelFindReplaceModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Modified_Document";

        [Required(ErrorMessage = "Search text is required")]
        public string SearchText { get; set; } = "Technology";

        public string ReplaceText { get; set; } = "New Text";

        [Required(ErrorMessage = "Operation type is required")]
        public FindReplaceActionType Operation { get; set; } = FindReplaceActionType.Replace;

        public string HighlightColor { get; set; }
        public bool MatchCase { get; set; } = false;
        public bool MatchEntireCellContent { get; set; } = false;
        internal object Color { get; set; } = "#FFFF00"; // Default yellow color

        internal readonly List<object> OperationTypes = new()
        {
            new { Text = "Replace Text", Value = FindReplaceActionType.Replace },
            new { Text = "Remove Text", Value = FindReplaceActionType.Remove },
            new { Text = "Highlight Text", Value = FindReplaceActionType.Highlight }
        };
        internal bool IsReplaceOperation => Operation == FindReplaceActionType.Replace;
        internal bool IsHighlightOperation => Operation == FindReplaceActionType.Highlight;

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();
            // Validate based on operation type
            switch(Operation) {
                case FindReplaceActionType.Replace:
                    if(string.IsNullOrWhiteSpace(ReplaceText)) {
                        results.Add(new ValidationResult(
                            "Replace text is required for replace operation",
                            new[] { nameof(ReplaceText) }));
                    }
                    break;

                case FindReplaceActionType.Highlight:
                    if(string.IsNullOrWhiteSpace(Color.ToString())) {
                        results.Add(new ValidationResult(
                            "Highlight color is required for highlight operation",
                            new[] { nameof(HighlightColor) }));
                    }
                    break;
            }
            return results;
        }
    }
}
