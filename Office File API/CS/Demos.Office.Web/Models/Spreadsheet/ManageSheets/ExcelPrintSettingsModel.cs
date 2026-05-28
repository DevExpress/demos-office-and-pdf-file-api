using System.ComponentModel.DataAnnotations;
using DevExpress.Drawing.Printing;
using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelPrintSettingsModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Print_Document";

        // Paper settings
        public DXPaperKind PaperSize { get; set; } = DXPaperKind.Letter;
        public PageOrientation PageOrientation { get; set; } = PageOrientation.Portrait;   

        // Custom paper size (when PaperSize is "Custom")
        [Range(0.1, 22, ErrorMessage = "Paper width must be between 0.1 and 22 inches")]
        public float CustomPaperWidth { get; set; } = 8.5f; // inches
        [Range(0.1, 22, ErrorMessage = "Paper height must be between 0.1 and 22 inches")]
        public float CustomPaperHeight { get; set; } = 11.0f; // inches

        // Margins (in inches)
        [Range(0, 22, ErrorMessage = "Top margin must be between 0 and 22 inches")]
        public float TopMargin { get; set; } = 1.0f;
        [Range(0, 22, ErrorMessage = "Bottom margin must be between 0 and 22 inches")]
        public float BottomMargin { get; set; } = 1.0f;
        [Range(0, 22, ErrorMessage = "Left margin must be between 0 and 22 inches")]
        public float LeftMargin { get; set; } = 1.0f;
        [Range(0, 22, ErrorMessage = "Right margin must be between 0 and 22 inches")]
        public float RightMargin { get; set; } = 1.0f;

        // Header and Footer offsets (in inches)
        [Range(0, 22, ErrorMessage = "Header offset must be between 0 and 22 inches")]
        public float HeaderOffset { get; set; } = 0.5f;
        [Range(0, 22, ErrorMessage = "Footer offset must be between 0 and 22 inches")]
        public float FooterOffset { get; set; } = 0.5f;

        // Header and Footer text
        public bool DifferentFirstPage { get; set; } = true;

        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string HeaderLeftSectionText { get; set; } = "Header Left";
        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string HeaderCenterSectionText { get; set; } = "Header Center";
        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string HeaderRightSectionText { get; set; } = "Header Right";

        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string FooterLeftSectionText { get; set; } = "Footer Left";
        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string FooterCenterSectionText { get; set; } = "Footer Center";
        [MaxLength(50, ErrorMessage = "Section Text cannot exceed 50 characters")]
        public string FooterRightSectionText { get; set; } = "Footer Right";

        // Print settings
        public bool FitToPage { get; set; } = true;

        [Range(1, 10, ErrorMessage = "Fit To Width must be between 1 and 10")]
        public int FitToWidth { get; set; } = 1;
        [Range(1, 10, ErrorMessage = "Fit To Height must be between 1 and 10")]
        public int FitToHeight { get; set; } = 1;
        [Range(0.05, 4, ErrorMessage = "Scale must be between 5% and 400%")]
        public float Scale { get; set; } = 1.0f;

        public bool PrintGridlines { get; set; } = false;
        public bool PrintHeadings { get; set; } = false;
        public bool BlackAndWhite { get; set; } = false;

        [Range(1, 10, ErrorMessage = "Start Page must be between 1 and 10")]
        public int StartPage { get; set; } = 1;

        public PageOrder PageOrder { get; set; } = PageOrder.DownThenOver;        

        internal readonly List<object> PaperSizes = new() {
            new { Text = "Letter (8.5\" × 11\")", Value = DXPaperKind.Letter },
            new { Text = "A4 (210mm × 297mm)", Value = DXPaperKind.A4 },
            new { Text = "Legal (8.5\" × 14\")", Value = DXPaperKind.Legal },
            new { Text = "A3 (297mm × 420mm)", Value = DXPaperKind.A3 },
            new { Text = "A5 (148mm × 210mm)", Value = DXPaperKind.A5 },
            new { Text = "Tabloid (11\" × 17\")", Value = DXPaperKind.Tabloid },
            new { Text = "Executive (7.25\" × 10.5\")", Value = DXPaperKind.Executive },
            new { Text = "Custom Size", Value = DXPaperKind.Custom }
        };

        internal readonly List<object> PageOrientations = new() {
            new { Text = "Portrait", Value = PageOrientation.Portrait },
            new { Text = "Landscape", Value = PageOrientation.Landscape }
        };

        internal readonly List<object> PageOrders = new() {
            new { Text = "Down, then over", Value = PageOrder.DownThenOver },
            new { Text = "Over, then down", Value = PageOrder.OverThenDown }
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            // Validate margins don't exceed reasonable bounds
            var maxMargin = GetMaxMarginForPaperSize();
            if(TopMargin + BottomMargin >= maxMargin.Height) {
                results.Add(new ValidationResult(
                    "Top and bottom margins combined exceed paper height",
                    new[] { nameof(TopMargin), nameof(BottomMargin), nameof(CustomPaperHeight) }));
            }

            if(LeftMargin + RightMargin >= maxMargin.Width) {
                results.Add(new ValidationResult(
                    "Left and right margins combined exceed paper width",
                    new[] { nameof(LeftMargin), nameof(RightMargin), nameof(CustomPaperWidth) }));
            }

            return results;
        }

        // Helper methods
        public bool IsCustomPaperSize() {
            return PaperSize == DXPaperKind.Custom;
        }

        (double Width, double Height) GetMaxMarginForPaperSize() {
            return PaperSize switch {
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
