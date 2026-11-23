using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.SearchExtract {
    public class ExcelMetadataModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Document_Metadata";

        public bool ReadDocumentProperties { get; set; } = true;
        public bool ReadCustomProperties { get; set; } = true;

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".xlt" };
    }
}
