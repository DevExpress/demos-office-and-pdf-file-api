using System.ComponentModel.DataAnnotations;
using DevExpress.Docs.Office;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentElements {
    public class PresentationChartsModel : PresentationModelBase {
        public override string OutputName { get; set; } = "Presentation_Charts";

        // Chart generation settings
        [Required(ErrorMessage = "Chart title is required")]
        [MaxLength(50, ErrorMessage = "Chart title must not exceed 50 characters")]
        public string ChartTitle { get; set; } = "Example Chart";

        [Range(3.0, 6.0, ErrorMessage = "Chart height must be between 3 and 6 inches")]
        public float ChartHeight { get; set; } = 5f; // inches

        public bool ShowTrendline { get; set; } = true;

        // Chart appearance settings
        public bool ShowLegend { get; set; } = true;
        public bool ShowDataLabels { get; set; } = true;
        public bool ShowGridLines { get; set; } = true;
        public ChartStyle ChartStyle { get; set; } = ChartStyle.Color;

        internal readonly List<object> ChartStyles = new() {
            new { Text = "Color", Value = ChartStyle.Color },
            new { Text = "Accent 3 Gradient", Value = ChartStyle.Accent3Gradient },
            new { Text = "Gray", Value = ChartStyle.Gray },
        };

        // Axis settings
        public bool ShowAxisTitles { get; set; } = true;

        // Legend settings
        [Required(ErrorMessage = "Legend position is required")]
        public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;

        internal readonly List<object> LegendPositions = new() {
            new { Text = "Top", Value = LegendPositionType.Top },
            new { Text = "Bottom", Value = LegendPositionType.Bottom },
            new { Text = "Left", Value = LegendPositionType.Left },
            new { Text = "Right", Value = LegendPositionType.Right },
            new { Text = "TopRight", Value = LegendPositionType.TopRight }
        };

        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo { Text = "PowerPoint Presentation (.pptx)", Value = "pptx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Presentation (.pptm)", Value = "pptm" },
            new FormatInfo { Text = "PowerPoint Template (.potx)", Value = "potx" },
            new FormatInfo { Text = "PowerPoint Macro-Enabled Template (.potm)", Value = "potm" },
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" }
        };

        public bool IsPdfOutput => OutputFormat?.Equals("pdf", StringComparison.OrdinalIgnoreCase) == true;
    }
}
