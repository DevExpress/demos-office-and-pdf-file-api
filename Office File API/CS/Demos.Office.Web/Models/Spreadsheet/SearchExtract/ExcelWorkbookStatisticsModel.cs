using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.SearchExtract {
    public class ExcelWorkbookStatisticsModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Workbook_Statistics";
        public bool SkipHiddenWorksheets { get; set; } = false;
    }
}
