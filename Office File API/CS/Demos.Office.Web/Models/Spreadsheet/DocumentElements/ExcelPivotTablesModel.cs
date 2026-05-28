using System.ComponentModel.DataAnnotations;
using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelPivotTablesModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "PivotTable_Document";

        // Pivot Table generation settings
        public bool CreatePivotTableSheet { get; set; } = false;

        // Layout settings
        [Required(ErrorMessage = "Report layout is required")]
        public PivotReportLayout ReportLayout { get; set; } = PivotReportLayout.Compact;

        internal readonly List<object> ReportLayoutTypes = new()
        {
            new { Text = "Compact", Value = PivotReportLayout.Compact },
            new { Text = "Outline", Value = PivotReportLayout.Outline },
            new { Text = "Tabular", Value = PivotReportLayout.Tabular }
        };

        public bool ShowRowGrandTotals { get; set; } = true;
        public bool ShowColumnGrandTotals { get; set; } = true;
        public bool ShowSubtotals { get; set; } = true;

        // Appearance and view settings
        public bool ShowRowHeaders { get; set; } = true;
        public bool ShowColumnHeaders { get; set; } = true;
        public bool ShowFieldHeaders { get; set; } = true;

        // Filter
        [Required(ErrorMessage = "Sales Region is required")]
        public string SalesRegion { get; set; } = "East";

        internal readonly List<object> SalesRegions = new()
        {
            new { Text = "East", Value = "East" },
            new { Text = "West", Value = "West" },
            new { Text = "South", Value = "South" },
            new { Text = "North", Value = "North" }
        };
    }
}
