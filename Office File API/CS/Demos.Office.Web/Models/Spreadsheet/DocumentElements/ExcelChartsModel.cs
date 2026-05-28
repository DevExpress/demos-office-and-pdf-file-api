using DevExpress.Spreadsheet.Charts;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelChartsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Chart_Document";

        // Chart generation settings
        public bool LoadFromTemplate { get; set; } = false;
        public bool CreateChartSheet { get; set; } = true;
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
    }
}
