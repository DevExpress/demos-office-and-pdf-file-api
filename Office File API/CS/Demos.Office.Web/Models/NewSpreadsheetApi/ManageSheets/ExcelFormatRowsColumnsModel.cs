using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets {
    public class ExcelFormatRowsColumnsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Modified_Document";

        public FormatOperationType OperationType { get; set; } = FormatOperationType.Add;
        [Range(1, 1048576)]
        public int RowCount { get; set; } = 1;
        [Range(1, 16384)]
        public int ColumnCount { get; set; } = 1;
        [Range(1, 1048576)]
        public int StartRowIndex { get; set; } = 1;
        [Range(1, 16384)]
        public int StartColumnIndex { get; set; } = 1;
        [Range(1, 1700)]
        public double RowHeight { get; set; } = 150;
        [Range(1, 250)]
        public double ColumnWidth { get; set; } = 20;

        internal string FormatTarget { get; set; } = "Rows";
        internal IEnumerable<string> FormatTargetOptions = new[] {
            "Rows",
            "Columns",
        };

        internal readonly List<object> OperationTypes = new()
        {
            new { Text = "Add", Value = FormatOperationType.Add },
            new { Text = "Remove", Value = FormatOperationType.Remove },
            new { Text = "Set Size", Value = FormatOperationType.SetSize },
            new { Text = "Hide", Value = FormatOperationType.Hide },
            new { Text = "AutoFit", Value = FormatOperationType.AutoFit },
        };
    }
    public enum FormatOperationType {
        Add,
        Remove,
        SetSize,
        Hide,
        AutoFit
    }
}
