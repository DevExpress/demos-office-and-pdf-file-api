using System.ComponentModel.DataAnnotations;
using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations {
    public class ExcelCalculateWorkbookModel : ExcelFileModelBase {
        // Override OutputFormat validation to include Excel formats
        [RegularExpression(@"^(pdf|xlsx|xls|xlsm|xltx|xlt|xltm|xlsb)$",
            ErrorMessage = "Invalid output format. Supported formats: pdf, xlsx, xls, xlsm, xltx, xlt, xltm, xlsb")]
        public override string OutputFormat { get; set; } = "pdf";
        public override string OutputName { get; set; } = "Calculated_Document";

        [RegularExpression(@"^(Automatic|AutomaticExceptTables|Manual)$",
            ErrorMessage = "Invalid value. Supported values: Automatic, AutomaticExceptTables, Manual")]
        public CalculationMode? DocumentCalculationMode { get; set; } = null;

        internal readonly List<object> DocumentCalculationModes = new() {
            new { Text = "Automatic", Value = CalculationMode.Automatic },
            new { Text = "Automatic Except Tables", Value = CalculationMode.AutomaticExceptTables },
            new { Text = "Manual", Value = CalculationMode.Manual },
        };
        public override List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo() { Text = "PDF Document (*.pdf)", Value = "pdf" },
            new FormatInfo() { Text = "Excel Workbook (*.xlsx)", Value = "xlsx" },
            new FormatInfo() { Text = "Excel 97-2003 Workbook (*.xls)", Value = "xls" },
            new FormatInfo() { Text = "Excel Macro-Enabled Workbook (*.xlsm)", Value = "xlsm" },
            new FormatInfo() { Text = "Excel Template (*.xltx)", Value = "xltx" },
            new FormatInfo() { Text = "Excel 97-2003 Template (*.xlt)", Value = "xlt" },
            new FormatInfo() { Text = "Excel Macro-Enabled Template (*.xltm)", Value = "xltm" },
            new FormatInfo() { Text = "Excel Binary Workbook (*.xlsb)", Value = "xlsb" },
        };

        public bool IsPdfOutput => OutputFormat?.Equals("pdf", StringComparison.OrdinalIgnoreCase) == true;
    }
}
