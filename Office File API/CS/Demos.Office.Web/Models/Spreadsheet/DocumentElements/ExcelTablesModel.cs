using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelTablesModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Table_Document";

        // Table structure
        [Range(1, 10, ErrorMessage = "Number of rows must be between 1 and 10")]
        public int NumberOfRows { get; set; } = 5;

        [Range(1, 10, ErrorMessage = "Number of columns must be between 1 and 10")]
        public int NumberOfColumns { get; set; } = 5;

        // Table options
        public bool CreateCustomTableStyle { get; set; } = false;
        public bool ShowTotals { get; set; } = true;
    }
}
