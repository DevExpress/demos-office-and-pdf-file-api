using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelGroupDataModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Group_Document";
        [RegularExpression(@"^(Rows|Columns|Both)$",
            ErrorMessage = "Invalid clear mode. Supported modes: Rows, Columns, Both")]
        public GroupMode GroupMode { get; set; } = GroupMode.Rows;

        public bool CollapseGroups { get; set; } = false;
        public bool ShowSubtotals { get; set; } = true;

        internal readonly List<object> GroupModes = new()
        {
            new { Text = "Rows", Value = GroupMode.Rows },
            new { Text = "Columns", Value = GroupMode.Columns },
            new { Text = "Both", Value = GroupMode.Both }
        };
    }
    public enum GroupMode {
        Rows,
        Columns,
        Both
    }
}
