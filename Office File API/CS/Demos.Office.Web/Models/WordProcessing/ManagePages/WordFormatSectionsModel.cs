using DevExpress.Drawing.Printing;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.ManagePages
{
    public class WordFormatSectionsModel : WordFileModelBase, IValidatableObject
    {
        // Paper settings
        [Required(ErrorMessage = "Paper size is required")]
        public DXPaperKind PaperSize { get; set; } = DXPaperKind.Letter;

        [Required(ErrorMessage = "Page orientation is required")]
        public string PageOrientation { get; set; } = "Portrait"; // Portrait, Landscape
        public bool IsLandscape => PageOrientation.ToLowerInvariant() == "landscape";

        // Custom paper size (when PaperSize is "Custom")
        [Range(0.1, 22, ErrorMessage = "Paper width must be between 0.1 and 22 inches")]
        public float CustomPaperWidth { get; set; } = 8.5f; // inches

        [Range(0.1, 22, ErrorMessage = "Paper height must be between 0.1 and 22 inches")]
        public float CustomPaperHeight { get; set; } = 11.0f; // inches

        // Margins (in inches)
        [Range(-22, 22, ErrorMessage = "Top margin must be between -22 and 22 inches")]
        public float TopMargin { get; set; } = 1.0f;

        [Range(-22, 22, ErrorMessage = "Bottom margin must be between -22 and 22 inches")]
        public float BottomMargin { get; set; } = 1.0f;

        [Range(0, 22, ErrorMessage = "Left margin must be between -22 and 22 inches")]
        public float LeftMargin { get; set; } = 1.0f;

        [Range(0, 22, ErrorMessage = "Right margin must be between -22 and 22 inches")]
        public float RightMargin { get; set; } = 1.0f;

        // Mirrored margins
        public bool MirroredMargins { get; set; } = false;

        // Header and Footer offsets (in inches)
        [Range(0, 22, ErrorMessage = "Header offset must be between 0 and 22 inches")]
        public float HeaderOffset { get; set; } = 0.5f;

        [Range(0, 22, ErrorMessage = "Footer offset must be between 0 and 22 inches")]
        public float FooterOffset { get; set; } = 0.5f;

        // Gutter settings

        [Range(0, 22, ErrorMessage = "Gutter size must be between 0 and 22 inches")]
        public float GutterSize { get; set; } = 0f;

        [Required(ErrorMessage = "Gutter position is required")]
        public string GutterPosition { get; set; } = "Left";

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "FormattedDocument";

        internal readonly List<object> PaperSizes = new()
        {
            new { Text = "Letter (8.5\" × 11\")", Value = DXPaperKind.Letter },
            new { Text = "A4 (210mm × 297mm)", Value = DXPaperKind.A4 },
            new { Text = "Legal (8.5\" × 14\")", Value = DXPaperKind.Legal },
            new { Text = "A3 (297mm × 420mm)", Value = DXPaperKind.A3 },
            new { Text = "A5 (148mm × 210mm)", Value = DXPaperKind.A5 },
            new { Text = "Tabloid (11\" × 17\")", Value = DXPaperKind.Tabloid },
            new { Text = "Executive (7.25\" × 10.5\")", Value = DXPaperKind.Executive },
            new { Text = "Custom Size", Value = DXPaperKind.Custom }
        };

        internal readonly List<object> PageOrientations = new()
        {
            new { Text = "Portrait", Value = "Portrait" },
            new { Text = "Landscape", Value = "Landscape" }
        };

        internal readonly List<object> GutterPositions = new()
        {
            new { Text = "Left", Value = "Left" },
            new { Text = "Right", Value = "Right" },
            new { Text = "Top", Value = "Top" }
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

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();

            // Validate custom paper size when Custom is selected
            if (PaperSize == DXPaperKind.Custom)
            {
                if (CustomPaperWidth <= 0)
                {
                    results.Add(new ValidationResult(
                        "Custom paper width must be greater than 0",
                        new[] { nameof(CustomPaperWidth) }));
                }

                if (CustomPaperHeight <= 0)
                {
                    results.Add(new ValidationResult(
                        "Custom paper height must be greater than 0",
                        new[] { nameof(CustomPaperHeight) }));
                }
            }

            // Validate gutter settings
            if (GutterSize < 0)
            {
                results.Add(new ValidationResult(
                    "Gutter size must be greater than 0 when gutter is enabled",
                    new[] { nameof(GutterSize) }));
            }

            var validPositions = new[] { "Left", "Right", "Top" };
            if (!validPositions.Contains(GutterPosition))
            {
                results.Add(new ValidationResult(
                    "Invalid gutter position",
                    new[] { nameof(GutterPosition) }));
            }

            // Validate margins don't exceed reasonable bounds
            var maxMargin = GetMaxMarginForPaperSize();
            if (TopMargin + BottomMargin >= maxMargin.Height)
            {
                results.Add(new ValidationResult(
                    "Top and bottom margins combined exceed paper height",
                    new[] { nameof(TopMargin), nameof(BottomMargin) }));
            }

            if (LeftMargin + RightMargin >= maxMargin.Width)
            {
                results.Add(new ValidationResult(
                    "Left and right margins combined exceed paper width",
                    new[] { nameof(LeftMargin), nameof(RightMargin) }));
            }

            return results;
        }

        // Helper methods
        public bool IsCustomPaperSize()
        {
            return PaperSize == DXPaperKind.Custom;
        }

        (double Width, double Height) GetMaxMarginForPaperSize()
        {
            return PaperSize switch
            {
                DXPaperKind.Letter => (8.5, 11.0),
                DXPaperKind.A4 => (8.27, 11.69), // 210mm x 297mm converted to inches
                DXPaperKind.Legal => (8.5, 14.0),
                DXPaperKind.A3 => (11.69, 16.54), // 297mm x 420mm converted to inches
                DXPaperKind.A5 => (5.83, 8.27), // 148mm x 210mm converted to inches
                DXPaperKind.Tabloid => (11.0, 17.0),
                DXPaperKind.Executive => (7.25, 10.5),
                DXPaperKind.Custom => (CustomPaperWidth, CustomPaperHeight),
                _ => (8.5, 11.0) // Default to Letter
            };
        }
    }
}
