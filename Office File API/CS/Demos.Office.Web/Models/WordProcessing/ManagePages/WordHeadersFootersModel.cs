using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages
{
    public class WordHeadersFootersModel : WordFileModelBase, IValidatableObject
    {
        // Header settings
        public bool AddHeader { get; set; } = false;
        public string HeaderText { get; set; } = "Document Title";
        public ParagraphAlignment HeaderAlignment { get; set; } = ParagraphAlignment.Center;

        // Footer settings
        public bool AddFooter { get; set; } = false;
        public string FooterText { get; set; } = "Confidential";
        public ParagraphAlignment FooterAlignment { get; set; } = ParagraphAlignment.Center;

        // Different first page settings
        public bool DifferentFirstPage { get; set; } = false;
        public string FirstPageHeaderText { get; set; } = "";
        public string FirstPageFooterText { get; set; } = "";
        public ParagraphAlignment FirstPageHeaderAlignment { get; set; } = ParagraphAlignment.Center;
        public ParagraphAlignment FirstPageFooterAlignment { get; set; } = ParagraphAlignment.Center;

        // Page numbering
        public bool AddPageNumbering { get; set; } = false;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Modified_Document";

        // Static data for UI components
        internal readonly List<object> AlignmentOptions = new()
        {
            new { Text = "Left", Value = ParagraphAlignment.Left },
            new { Text = "Center", Value = ParagraphAlignment.Center },
            new { Text = "Right", Value = ParagraphAlignment.Right }
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

        // Implement IValidatableObject for custom validation
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();

            // Validate header settings
            if (AddHeader && string.IsNullOrWhiteSpace(HeaderText))
            {
                results.Add(new ValidationResult(
                    "Header text is required when adding a header",
                    new[] { nameof(HeaderText) }));
            }

            // Validate footer settings
            if (AddFooter && string.IsNullOrWhiteSpace(FooterText))
            {
                results.Add(new ValidationResult(
                    "Footer text is required when adding a footer",
                    new[] { nameof(FooterText) }));
            }

            // Validate alignment values
            var validAlignments = new[] { ParagraphAlignment.Left, ParagraphAlignment.Center, ParagraphAlignment.Right };
            if (!validAlignments.Contains(HeaderAlignment))
            {
                results.Add(new ValidationResult(
                    "Invalid header alignment",
                    new[] { nameof(HeaderAlignment) }));
            }

            if (!validAlignments.Contains(FooterAlignment))
            {
                results.Add(new ValidationResult(
                    "Invalid footer alignment",
                    new[] { nameof(FooterAlignment) }));
            }

            return results;
        }

        // Helper methods for UI logic
        public bool IsPageNumberingAvailable()
        {
            // Page numbering is only available if at least one of header or footer is enabled
            return AddHeader || AddFooter;
        }

        public List<object> GetAvailablePageNumberPositions()
        {
            var availablePositions = new List<object>();

            // Only add header option if header is enabled
            if (AddHeader)
            {
                availablePositions.Add(new { Text = "Header", Value = "header" });
            }

            // Only add footer option if footer is enabled
            if (AddFooter)
            {
                availablePositions.Add(new { Text = "Footer", Value = "footer" });
            }

            return availablePositions;
        }
    }
}
