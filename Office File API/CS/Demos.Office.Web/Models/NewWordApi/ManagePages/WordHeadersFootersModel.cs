using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Word;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.ManagePages {
    public class WordHeadersFootersModel : WordFileModelBase, IValidatableObject {
        public bool AddHeader { get; set; } = false;
        public string HeaderText { get; set; } = "Document Title";
        public ParagraphAlignment HeaderAlignment { get; set; } = ParagraphAlignment.Center;

        public bool AddFooter { get; set; } = false;
        public string FooterText { get; set; } = "Confidential";
        public ParagraphAlignment FooterAlignment { get; set; } = ParagraphAlignment.Center;

        public bool DifferentFirstPage { get; set; } = false;
        public string FirstPageHeaderText { get; set; } = "";
        public string FirstPageFooterText { get; set; } = "";
        public ParagraphAlignment FirstPageHeaderAlignment { get; set; } = ParagraphAlignment.Center;
        public ParagraphAlignment FirstPageFooterAlignment { get; set; } = ParagraphAlignment.Center;

        public bool AddPageNumbering { get; set; } = false;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Modified_Document";

        internal readonly List<object> AlignmentOptions = new() {
            new { Text = "Left", Value = ParagraphAlignment.Left },
            new { Text = "Center", Value = ParagraphAlignment.Center },
            new { Text = "Right", Value = ParagraphAlignment.Right }
        };

        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            if (AddHeader && string.IsNullOrWhiteSpace(HeaderText)) {
                results.Add(new ValidationResult(
                    "Header text is required when adding a header",
                    new[] { nameof(HeaderText) }));
            }

            if (AddFooter && string.IsNullOrWhiteSpace(FooterText)) {
                results.Add(new ValidationResult(
                    "Footer text is required when adding a footer",
                    new[] { nameof(FooterText) }));
            }

            var validAlignments = new[] { ParagraphAlignment.Left, ParagraphAlignment.Center, ParagraphAlignment.Right };
            if (!validAlignments.Contains(HeaderAlignment)) {
                results.Add(new ValidationResult(
                    "Invalid header alignment",
                    new[] { nameof(HeaderAlignment) }));
            }

            if (!validAlignments.Contains(FooterAlignment)) {
                results.Add(new ValidationResult(
                    "Invalid footer alignment",
                    new[] { nameof(FooterAlignment) }));
            }

            return results;
        }
    }
}