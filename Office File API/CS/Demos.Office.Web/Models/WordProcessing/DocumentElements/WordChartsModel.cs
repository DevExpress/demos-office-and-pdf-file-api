using DevExpress.Spreadsheet.Charts;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements {
    public class WordChartsModel : WordFileModelBase {
        // Override to support more file types for chart operations
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Chart_Document";

        // Chart generation settings
        public bool LoadFromTemplate { get; set; } = false;
        public bool ShowTrendline { get; set; } = true;

        // Chart style settings
        public bool ShowLegend { get; set; } = true;
        public bool ShowDataLabels { get; set; } = true;
        public bool ShowGridLines { get; set; } = true;

        [Required(ErrorMessage = "Legend position is required")]
        public LegendPosition LegendPosition { get; set; } = LegendPosition.Right;

        // Axis settings
        public bool ShowAxisTitles { get; set; } = true;

        internal readonly List<object> LegendPositions = new()
        {
            new { Text = "Top", Value = LegendPosition.Top },
            new { Text = "Bottom", Value = LegendPosition.Bottom },
            new { Text = "Left", Value = LegendPosition.Left },
            new { Text = "Right", Value = LegendPosition.Right }
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
    }
}
