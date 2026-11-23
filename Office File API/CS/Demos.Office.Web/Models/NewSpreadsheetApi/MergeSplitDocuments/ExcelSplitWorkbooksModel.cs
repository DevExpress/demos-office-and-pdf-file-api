using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.MergeSplitDocuments {
    public class ExcelSplitWorkbooksModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Split_Documents";
        [RegularExpression(@"^(VisibleOnly|All)$",
            ErrorMessage = "Invalid visibility mode. Supported modes: VisibleOnly, All")]
        public VisibilityMode SplitMode { get; set; } = VisibilityMode.VisibleOnly;

        internal readonly List<object> SplitModes = new() {
            new { Text = "Visible Only", Value = VisibilityMode.VisibleOnly },
            new { Text = "All", Value = VisibilityMode.All },
        };
    }

    public enum VisibilityMode {
        VisibleOnly,
        All
    }
}
