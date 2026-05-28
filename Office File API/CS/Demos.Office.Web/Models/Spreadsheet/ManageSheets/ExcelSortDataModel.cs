using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelSortDataModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Sorted_Data";

        [RegularExpression(@"^(SingleColumn|MultipleColumns|FontColor|Fill|Custom)$",
            ErrorMessage = "Invalid clear mode. Supported modes: SingleColumn, MultipleColumns, FontColor, Fill, Custom")]
        public SortMode SortMode { get; set; } = SortMode.SingleColumn;

        internal readonly List<object> SortModes = new() {
            new { Text = "Single Column", Value = SortMode.SingleColumn },
            new { Text = "Multiple Columns", Value = SortMode.MultipleColumns },
            new { Text = "Font Color", Value = SortMode.FontColor },
            new { Text = "Fill Color", Value = SortMode.Fill },
            new { Text = "Custom", Value = SortMode.Custom }
        };

        public string SingleSortColumnName { get; set; } = "Continent";

        public IEnumerable<string> MultipleSortColumnNames { get; set; } = new List<string>{
            "Continent",
            "Country",
        };

        internal readonly List<string> ColumnNames = new() {
            "Continent",
            "Country",
            "Symbol",
            "Population",
            "GDP",
            "Area"
        };

        public SortOrder SortOrder { get; set; } = SortOrder.Ascending;

        internal readonly List<SortOrder> SortOrderOptions = new() {
              SortOrder.Ascending,
              SortOrder.Descending
        };
    }
    public enum SortMode {
        SingleColumn,
        MultipleColumns,
        FontColor,
        Fill,
        Custom
    }
    public enum SortOrder {
        Ascending,
        Descending
    }
}
