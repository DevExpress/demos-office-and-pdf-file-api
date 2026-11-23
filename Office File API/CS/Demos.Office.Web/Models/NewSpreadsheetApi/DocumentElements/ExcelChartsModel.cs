using System.ComponentModel.DataAnnotations;
using DevExpress.Docs.Office;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentElements {
    public class ExcelChartsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Chart_Document";

        public bool CreateChartSheet { get; set; } = true;
        public bool ShowTrendline { get; set; } = true;

        public bool ShowLegend { get; set; } = true;
        public bool ShowDataLabels { get; set; } = true;
        public bool ShowGridLines { get; set; } = true;

        [Required(ErrorMessage = "Legend position is required")]
        public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;

        public bool ShowAxisTitles { get; set; } = true;

        internal readonly List<object> LegendPositions = new()
        {
            new { Text = "Top", Value = LegendPositionType.Top },
            new { Text = "Bottom", Value = LegendPositionType.Bottom },
            new { Text = "Left", Value = LegendPositionType.Left },
            new { Text = "Right", Value = LegendPositionType.Right }
        };

        public override List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo { Text = "Excel Workbook (*.xlsx)", Value = "xlsx" },
            new FormatInfo { Text = "Excel Macro-Enabled Workbook (*.xlsm)", Value = "xlsm" },
            new FormatInfo { Text = "Excel Template (*.xltx)", Value = "xltx" },
            new FormatInfo { Text = "Excel Macro-Enabled Template (*.xltm)", Value = "xltm" },
            new FormatInfo { Text = "Excel Binary Workbook (*.xlsb)", Value = "xlsb" }
        };
    }
}
