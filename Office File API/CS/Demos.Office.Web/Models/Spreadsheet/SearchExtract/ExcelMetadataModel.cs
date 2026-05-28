using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.SearchExtract {
    public class ExcelMetadataModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Document_Metadata";

        public bool ReadDocumentProperties { get; set; } = true;
        public bool ReadCustomProperties { get; set; } = true;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt" };
    }
}
