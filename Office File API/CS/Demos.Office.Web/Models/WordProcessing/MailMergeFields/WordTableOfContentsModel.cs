using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.MailMergeFields
{
    public class WordTableOfContentsModel : WordFileModelBase, IValidatableObject
    {
        public override string OutputName { get; set; } = "Table_of_Contents";
        public bool UseHyperlinks { get; set; } = true;
        public bool IncludePageNumbers { get; set; } = true;
        [Range(1, 9, ErrorMessage = "Start Level must be between 1 and 9")]
        public int StartLevel { get; set; } = 1;
        [Range(1, 9, ErrorMessage = "End Level must be between 1 and 9")]
        public int EndLevel { get; set; } = 3;
        public TOCInsertMode InsertMode { get; set; } = TOCInsertMode.OutlineLevels;
        public TabLeaderType TabLeaderType { get; set; } = TabLeaderType.Dots;

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal readonly List<object> TabLeaderTypes = new()
        {
            new { Text = "None", Value = TabLeaderType.None },
            new { Text = "Underscores (___)", Value = TabLeaderType.Underline },
            new { Text = "Hyphens (---)", Value = TabLeaderType.Hyphens },
            new { Text = "Dots (...)", Value = TabLeaderType.Dots }
        };
        internal readonly List<object> InsertModes = new()
        {
            new { Text = "Outline Levels", Value = TOCInsertMode.OutlineLevels },
            new { Text = "Paragraph Styles", Value = TOCInsertMode.ParagraphStyles },
            new { Text = "TC Entries", Value = TOCInsertMode.TCFields }
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (EndLevel < StartLevel)
            {
                yield return new ValidationResult("End Level must be greater than Start Level", new[] { nameof(EndLevel) });
            }
        }
    }
    public enum TOCInsertMode
    {
        OutlineLevels,
        ParagraphStyles,
        TCFields
    }
}
