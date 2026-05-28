using System.ComponentModel.DataAnnotations;
using DevExpress.XtraRichEdit.API.Native;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.ReviewComment
{
    public class WordCompareDocumentsModel : WordFileModelBase
    {
        public bool CaseChanges { get; set; } = true;

        public bool Formatting { get; set; } = true;

        public bool HeadersFooters { get; set; } = true;

        public bool TextBoxes { get; set; } = true;

        public string Author { get; set; } = "Test Author";

        [Required]
        [EnumDataType(typeof(ComparisonLevel))]
        public ComparisonLevel ComparisonLevel { get; set; } = ComparisonLevel.Word;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf"};
        public override string OutputName { get; set; } = "CompareDocuments_Result";

        internal readonly IReadOnlyList<string> CompareOptions = new List<string> {
            "Case Changes",
            "Formatting",
            "Headers and Footers",
            "Textboxes"
        };

        internal IEnumerable<string> SelectedCompareOptions {
            get {
                if(CaseChanges)
                    yield return "Case Changes";
                if(Formatting)
                    yield return "Formatting";
                if(HeadersFooters)
                    yield return "Headers and Footers";
                if(TextBoxes)
                    yield return "Textboxes";
            }
            set {
                var selected = value?.ToHashSet() ?? new HashSet<string>();
                CaseChanges = selected.Contains("Case Changes");
                Formatting = selected.Contains("Formatting");
                HeadersFooters = selected.Contains("Headers and Footers");
                TextBoxes = selected.Contains("Textboxes");
            }
        }

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
            internal readonly List<object> ComparisonLevels = new()
        {
            new { Text = "Characters", Value = ComparisonLevel.Character },
            new { Text = "Words", Value = ComparisonLevel.Word }
        };

        }
}
