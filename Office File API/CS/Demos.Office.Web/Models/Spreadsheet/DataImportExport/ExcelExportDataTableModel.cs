namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DataImportExport {
    public class ExcelExportDataTableModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "json";
        public override string OutputName { get; set; } = "Exported_Data";
    }
}
