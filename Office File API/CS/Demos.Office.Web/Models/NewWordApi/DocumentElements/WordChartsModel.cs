using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Office;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements {
    public class WordChartsModel : WordFileModelBase {
        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Chart_Document";

        public bool ShowTrendline { get; set; } = true;

        public bool ShowLegend { get; set; } = true;
        public bool ShowDataLabels { get; set; } = true;
        public bool ShowGridLines { get; set; } = true;

        [Required(ErrorMessage = "Legend position is required")]
        public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;

        public bool ShowAxisTitles { get; set; } = true;

        internal readonly List<object> LegendPositions = new() {
            new { Text = "Top", Value = LegendPositionType.Top },
            new { Text = "Bottom", Value = LegendPositionType.Bottom },
            new { Text = "Left", Value = LegendPositionType.Left },
            new { Text = "Right", Value = LegendPositionType.Right }
        };

        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
    }
}