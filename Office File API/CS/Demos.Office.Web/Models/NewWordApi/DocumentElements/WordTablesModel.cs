using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Word;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements {
    public class WordTablesModel : WordFileModelBase, IValidatableObject {
        [Range(1, 50, ErrorMessage = "Number of rows must be between 1 and 50")]
        public int NumberOfRows { get; set; } = 5;

        [Range(1, 20, ErrorMessage = "Number of columns must be between 1 and 20")]
        public int NumberOfColumns { get; set; } = 4;

        [Required(ErrorMessage = "Table size mode is required")]
        public TableSizeModeType TableSizeMode { get; set; } = TableSizeModeType.AutoFitToWindow;

        [Range(0.1, 10.0, ErrorMessage = "Column width must be between 0.1 and 10.0 inches")]
        public float FixedColumnWidth { get; set; } = 1.5f;

        public TableBorderType BorderType { get; set; } = TableBorderType.AllBorders;
        [Required(ErrorMessage = "Border style is required")]
        public BorderLineStyle BorderStyle { get; set; } = BorderLineStyle.Single;

        [Range(0.25, 6.0, ErrorMessage = "Border width must be between 0.25 and 6.0 points")]
        public float BorderWidth { get; set; } = 1.0f;

        public string BorderColor { get; set; }
        internal object Color { get; set; } = "#000000";

        [Required(ErrorMessage = "Text wrapping is required")]
        public TableTextWrapping TextWrapping { get; set; } = TableTextWrapping.None;

        [Range(-22.0f, 22.0f, ErrorMessage = "Horizontal position must be between -22.0 and 22.0 inches")]
        public float HorizontalPosition { get; set; } = 0.0f;

        [Range(-22f, 22.0f, ErrorMessage = "Vertical position must be between -22.0 and 22.0 inches")]
        public float VerticalPosition { get; set; } = 0.0f;

        public bool IncludeHeaders { get; set; } = true;
        public bool IncludeSampleData { get; set; } = true;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Table_Result";

        internal readonly List<object> TableSizeModes = new() {
            new { Text = "AutoFit to Window", Value = TableSizeModeType.AutoFitToWindow },
            new { Text = "AutoFit to Contents", Value = TableSizeModeType.AutoFitToContents },
            new { Text = "Fixed Column Width", Value = TableSizeModeType.FixedColumnWidth }
        };

        internal readonly List<object> BorderStyles = new() {
            new { Text = "Single Line", Value = BorderLineStyle.Single },
            new { Text = "Double Line", Value = BorderLineStyle.Double },
            new { Text = "Thick Line", Value = BorderLineStyle.Thick },
            new { Text = "Dashed", Value = BorderLineStyle.Dashed },
            new { Text = "Dotted", Value = BorderLineStyle.Dotted },
            new { Text = "No Border", Value = BorderLineStyle.None }
        };
        internal readonly List<object> BorderTypes = new() {
            new { Text = "All Borders", Value = TableBorderType.AllBorders },
            new { Text = "Outside Borders", Value = TableBorderType.OutsideBorders },
            new { Text = "Inside Borders", Value = TableBorderType.InsideBorders },
            new { Text = "Inside Horizontal Borders", Value = TableBorderType.InsideHorizontalBorders },
            new { Text = "Inside Vertical Borders", Value = TableBorderType.InsideVerticalBorders },
            new { Text = "No Borders", Value = TableBorderType.NoBorders }
        };

        internal readonly List<object> TextWrappingOptions = new() {
            new { Text = "None", Value = TableTextWrapping.None },
            new { Text = "Around", Value = TableTextWrapping.Around }
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

            if (NumberOfRows <= 0 || NumberOfRows > 50) {
                results.Add(new ValidationResult(
                    "Number of rows must be between 1 and 50",
                    new[] { nameof(NumberOfRows) }));
            }

            if (NumberOfColumns <= 0 || NumberOfColumns > 20) {
                results.Add(new ValidationResult(
                    "Number of columns must be between 1 and 20",
                    new[] { nameof(NumberOfColumns) }));
            }

            if (TableSizeMode == TableSizeModeType.FixedColumnWidth) {
                if (FixedColumnWidth <= 0 || FixedColumnWidth > 10) {
                    results.Add(new ValidationResult(
                        "Fixed column width must be between 0.1 and 10.0 inches",
                        new[] { nameof(FixedColumnWidth) }));
                }
            }

            if (BorderWidth <= 0 || BorderWidth > 6) {
                results.Add(new ValidationResult(
                    "Border width must be between 0.25 and 6.0 points",
                    new[] { nameof(BorderWidth) }));
            }

            if (TextWrapping == TableTextWrapping.Around) {
                if (HorizontalPosition < -50 || HorizontalPosition > 50) {
                    results.Add(new ValidationResult(
                        "Horizontal position must be between -50.0 and 50.0 inches",
                        new[] { nameof(HorizontalPosition) }));
                }

                if (VerticalPosition < -50 || VerticalPosition > 50) {
                    results.Add(new ValidationResult(
                        "Vertical position must be between -50.0 and 50.0 inches",
                        new[] { nameof(VerticalPosition) }));
                }
            }

            return results;
        }

        internal bool IsFixedColumnWidth() {
            return TableSizeMode == TableSizeModeType.FixedColumnWidth;
        }

        internal bool IsAroundWrapping() {
            return TextWrapping == TableTextWrapping.Around;
        }

        internal bool IsBorderLineNone() {
            return BorderStyle == BorderLineStyle.None;
        }
        internal bool IsBorderNone() {
            return BorderType == TableBorderType.NoBorders;
        }
    }
    public enum TableSizeModeType {
        AutoFitToWindow,
        AutoFitToContents,
        FixedColumnWidth
    }
    public enum TableBorderType {
        NoBorders,
        AllBorders,
        OutsideBorders,
        InsideBorders,
        InsideHorizontalBorders,
        InsideVerticalBorders
    }
}