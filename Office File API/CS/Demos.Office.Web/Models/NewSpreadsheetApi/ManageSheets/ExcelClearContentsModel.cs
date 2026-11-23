using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets {
    public class ExcelClearContentsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Modified_Document";
        [RegularExpression(@"^(All|CellContent|CellFormatting|CellHyperlinks|CellComments)$",
            ErrorMessage = "Invalid clear mode. Supported modes: All, CellContent, CellFormatting, CellHyperlinks, CellComments")]
        public ClearMode ClearMode { get; set; } = ClearMode.All;

        internal readonly List<object> ClearModes = new()
        {
            new { Text = "All", Value = ClearMode.All },
            new { Text = "Cell Content", Value = ClearMode.CellContent },
            new { Text = "Cell Formatting", Value = ClearMode.CellFormatting },
            new { Text = "Cell Hyperlinks", Value = ClearMode.CellHyperlinks },
            new { Text = "Cell Comments", Value = ClearMode.CellComments },
        };
    }
    public enum ClearMode {
        All,
        CellContent,
        CellFormatting,
        CellHyperlinks,
        CellComments
    }
}
