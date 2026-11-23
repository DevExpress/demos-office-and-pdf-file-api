using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.SearchExtract {
    public class ExcelWorkbookStatisticsModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Workbook_Statistics";
        public bool SkipHiddenWorksheets { get; set; } = false;
    }
}
