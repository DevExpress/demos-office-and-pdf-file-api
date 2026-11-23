using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.MergeSplitDocuments {
    public class ExcelMergeWorkbooksModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Merged_Document";
    }
}
