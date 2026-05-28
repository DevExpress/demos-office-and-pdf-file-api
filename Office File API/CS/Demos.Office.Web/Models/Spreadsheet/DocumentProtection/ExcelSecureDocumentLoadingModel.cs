
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentProtection {
    public class ExcelSecureDocumentLoadingModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Secure_Document";

        // Loading limits
        [Range(1, 10, ErrorMessage = "Please enter a value between 1 and 10.")]
        public int MaxWorksheetCount { get; set; } = 1;

        [Range(1, 500, ErrorMessage = "Please enter a value between 1 and 500.")]
        public int MaxSheetRowCount { get; set; } = 100;

        [Range(1, 1000, ErrorMessage = "Please enter a value between 1 and 1000.")]
        public int MaxSheetCellCount { get; set; } = 500;

        [Range(1, 10, ErrorMessage = "Please enter a value between 1 and 10.")]
        public int MaxChartCount { get; set; } = 1;

        // Loading options
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveExternalConnections { get; set; } = true;
        public bool RemoveExternalWorkbooks { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public bool RemovePivotCaches { get; set; } = true;
        public bool RemoveRestrictedFormulas { get; set; } = true;
        public bool RemoveRestrictedLinks { get; set; } = true;
    }
}
