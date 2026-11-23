using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Office;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.SearchExtract {
    public class WordFindReplaceModel : WordFileModelBase, IValidatableObject {
        [Required(ErrorMessage = "Search text is required")]
        public string SearchText { get; set; } = "search text";

        public string ReplaceText { get; set; } = "replacement text";

        [Required(ErrorMessage = "Operation type is required")]
        public FindReplaceActionType Operation { get; set; } = FindReplaceActionType.Replace;

        public TextSearchOptions SearchOptions => new TextSearchOptions();

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf", ".html" };
        public override string OutputName { get; set; } = "Modified_Document";

        internal readonly List<object> OperationTypes = new() {
            new { Text = "Replace Text", Value = FindReplaceActionType.Replace },
            new { Text = "Remove Text", Value = FindReplaceActionType.Remove }
        };
        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal bool IsReplaceOperation => Operation == FindReplaceActionType.Replace;
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            if(Operation == FindReplaceActionType.Replace && string.IsNullOrWhiteSpace(ReplaceText)) {
                results.Add(new ValidationResult(
                    "Replace text is required for replace operation",
                    new[] { nameof(ReplaceText) }));
            }

            return results;
        }
    }
}