using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.SearchExtract {
    public class WordFindReplaceModel : WordFileModelBase, IValidatableObject {
        [Required(ErrorMessage = "Search text is required")]
        public string SearchText { get; set; } = "search text";

        public string ReplaceText { get; set; } = "replacement text";

        [Required(ErrorMessage = "Operation type is required")]
        public FindReplaceActionType Operation { get; set; } = FindReplaceActionType.Replace;

        public string HighlightColor { get; set; }
        public SearchOptions SearchOptions => SearchOptions.None;
        internal object Color { get; set; } = "#FFFF00"; // Default yellow color

        // Override to support more file types for find/replace operations
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf", ".html" };
        public override string OutputName { get; set; } = "Modified_Document";

        internal readonly List<object> OperationTypes = new()
    {
        new { Text = "Replace Text", Value = FindReplaceActionType.Replace },
        new { Text = "Remove Text", Value = FindReplaceActionType.Remove },
        new { Text = "Highlight Text", Value = FindReplaceActionType.Highlight }
    };
        internal readonly List<object> SearchOptionTypes = new()
    {
        new { Text = "None", Value = SearchOptions.None },
        new { Text = "Case Sensitive", Value = SearchOptions.CaseSensitive },
        new { Text = "Whole Word", Value = SearchOptions.WholeWord }
    };
        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
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
