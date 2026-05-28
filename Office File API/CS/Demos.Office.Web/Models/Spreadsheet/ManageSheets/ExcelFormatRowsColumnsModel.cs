using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
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
        [Range(1, 40)]
        public double RowHeight { get; set; } = 1;
        [Range(1, 40)]
        public double ColumnWidth { get; set; } = 1;

        internal string FormatTarget { get; set; } = "Rows";
        internal IEnumerable<string> FormatTargetOptions = new[] {
            "Rows",
            "Columns",
        };
        public string DocumentUnit { get; set; } = "Inch";

        internal readonly List<string> Units = new()
        {
            "Centimeter",
            "Document",
            "Inch",
            "Millimeter",
            "Point"
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
