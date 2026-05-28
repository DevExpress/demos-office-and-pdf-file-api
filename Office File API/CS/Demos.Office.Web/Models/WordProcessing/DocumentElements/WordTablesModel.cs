using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements
{
    public class WordTablesModel : WordFileModelBase, IValidatableObject
    {
        // Table structure
        [Range(1, 50, ErrorMessage = "Number of rows must be between 1 and 50")]
        public int NumberOfRows { get; set; } = 5;

        [Range(1, 20, ErrorMessage = "Number of columns must be between 1 and 20")]
        public int NumberOfColumns { get; set; } = 4;

        // Table size mode
        [Required(ErrorMessage = "Table size mode is required")]
        public AutoFitBehaviorType TableSizeMode { get; set; } = AutoFitBehaviorType.AutoFitToWindow;

        [Range(0.1, 10.0, ErrorMessage = "Column width must be between 0.1 and 10.0 inches")]
        public float FixedColumnWidth { get; set; } = 1.5f; // inches

        // Border settings
        public BorderType BorderType { get; set; } = BorderType.AllBorders;
        [Required(ErrorMessage = "Border style is required")]
        public BorderLineStyle BorderStyle { get; set; } = BorderLineStyle.Single;

        [Range(0.25, 6.0, ErrorMessage = "Border width must be between 0.25 and 6.0 points")]
        public float BorderWidth { get; set; } = 1.0f; // points

        public string BorderColor { get; set; }
        internal object Color { get; set; } = "#000000";

        // Text wrapping
        [Required(ErrorMessage = "Text wrapping is required")]
        public TableTextWrappingType TextWrapping { get; set; } = TableTextWrappingType.None;

        // Table position (when text wrapping is Around)
        [Range(-22.0f, 22.0f, ErrorMessage = "Horizontal position must be between -22.0 and 22.0 inches")]
        public float HorizontalPosition { get; set; } = 0.0f; // inches from left margin

        [Range(-22f, 22.0f, ErrorMessage = "Vertical position must be between -22.0 and 22.0 inches")]
        public float VerticalPosition { get; set; } = 0.0f; // inches from top margin

        // Table content options
        public bool IncludeHeaders { get; set; } = true;
        public bool IncludeSampleData { get; set; } = true;

        // Override to support more file types for table operations
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Table_Result";

        // Static data for UI components
        internal readonly List<object> TableSizeModes = new()
        {
            new { Text = "AutoFit to Window", Value = AutoFitBehaviorType.AutoFitToWindow },
            new { Text = "AutoFit to Contents", Value = AutoFitBehaviorType.AutoFitToContents },
            new { Text = "Fixed Column Width", Value = AutoFitBehaviorType.FixedColumnWidth }
        };

        internal readonly List<object> BorderStyles = new()
        {
            new { Text = "Single Line", Value = BorderLineStyle.Single },
            new { Text = "Double Line", Value = BorderLineStyle.Double },
            new { Text = "Thick Line", Value = BorderLineStyle.Thick },
            new { Text = "Dashed", Value = BorderLineStyle.Dashed },
            new { Text = "Dotted", Value = BorderLineStyle.Dotted },
            new { Text = "No Border", Value = BorderLineStyle.None }
        };
        internal readonly List<object> BorderTypes = new()
        {
            new { Text = "All Borders", Value = BorderType.AllBorders },
            new { Text = "Outside Borders", Value = BorderType.OutsideBorders },
            new { Text = "Inside Borders", Value = BorderType.InsideBorders },
            new { Text = "Inside Horizontal Borders", Value = BorderType.InsideHorizontalBorders },
            new { Text = "Inside Vertical Borders", Value = BorderType.InsideVerticalBorders },
            new { Text = "No Borders", Value = BorderType.NoBorders }
        };

        internal readonly List<object> TextWrappingOptions = new()
        {
            new { Text = "None", Value = TableTextWrappingType.None },
            new { Text = "Around", Value = TableTextWrappingType.Around }
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

            // Validate table dimensions
            if (NumberOfRows <= 0 || NumberOfRows > 50)
            {
                results.Add(new ValidationResult(
                    "Number of rows must be between 1 and 50",
                    new[] { nameof(NumberOfRows) }));
            }

            if (NumberOfColumns <= 0 || NumberOfColumns > 20)
            {
                results.Add(new ValidationResult(
                    "Number of columns must be between 1 and 20",
                    new[] { nameof(NumberOfColumns) }));
            }

            // Validate table size mode
            var validSizeModes = new[] { AutoFitBehaviorType.FixedColumnWidth, AutoFitBehaviorType.AutoFitToWindow, AutoFitBehaviorType.AutoFitToContents };
            if (!validSizeModes.Contains(TableSizeMode))
            {
                results.Add(new ValidationResult(
                    "Invalid table size mode",
                    new[] { nameof(TableSizeMode) }));
            }

            // Validate fixed column width when applicable
            if (TableSizeMode == AutoFitBehaviorType.FixedColumnWidth)
            {
                if (FixedColumnWidth <= 0 || FixedColumnWidth > 10)
                {
                    results.Add(new ValidationResult(
                        "Fixed column width must be between 0.1 and 10.0 inches",
                        new[] { nameof(FixedColumnWidth) }));
                }
            }

            // Validate border style
            var validBorderStyles = new[] { BorderLineStyle.Single, BorderLineStyle.Double, BorderLineStyle.Thick, BorderLineStyle.Dashed, BorderLineStyle.Dotted, BorderLineStyle.None };
            if (!validBorderStyles.Contains(BorderStyle))
            {
                results.Add(new ValidationResult(
                    "Invalid border style",
                    new[] { nameof(BorderStyle) }));
            }

            // Validate border width
            if (BorderWidth <= 0 || BorderWidth > 6)
            {
                results.Add(new ValidationResult(
                    "Border width must be between 0.25 and 6.0 points",
                    new[] { nameof(BorderWidth) }));
            }

            // Validate text wrapping
            var validWrappingOptions = new[] { TableTextWrappingType.None, TableTextWrappingType.Around };
            if (!validWrappingOptions.Contains(TextWrapping))
            {
                results.Add(new ValidationResult(
                    "Invalid text wrapping option",
                    new[] { nameof(TextWrapping) }));
            }

            // Validate position when text wrapping is Around
            if (TextWrapping == TableTextWrappingType.Around)
            {
                if (HorizontalPosition < -50 || HorizontalPosition > 50)
                {
                    results.Add(new ValidationResult(
                        "Horizontal position must be between -50.0 and 50.0 inches",
                        new[] { nameof(HorizontalPosition) }));
                }

                if (VerticalPosition < -50 || VerticalPosition > 50)
                {
                    results.Add(new ValidationResult(
                        "Vertical position must be between -50.0 and 50.0 inches",
                        new[] { nameof(VerticalPosition) }));
                }
            }

            return results;
        }

        // Helper methods
        internal bool IsFixedColumnWidth()
        {
            return TableSizeMode == AutoFitBehaviorType.FixedColumnWidth;
        }

        internal bool IsAroundWrapping()
        {
            return TextWrapping == TableTextWrappingType.Around;
        }

        internal bool IsBorderLineNone()
        {
            return BorderStyle == BorderLineStyle.None;
        }
        internal bool IsBorderNone()
        {
            return BorderType == BorderType.NoBorders;
        }
    }
    public enum BorderType
    {
        NoBorders,
        AllBorders,
        OutsideBorders,
        InsideBorders,
        InsideHorizontalBorders,
        InsideVerticalBorders
    }
}
