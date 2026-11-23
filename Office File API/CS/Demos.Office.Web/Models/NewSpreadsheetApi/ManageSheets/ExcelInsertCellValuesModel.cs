using BlazorDemo.OfficeFileApi.Services;
using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets {
    public class ExcelInsertCellValuesModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Cell_Values";

        [Required]
        [MinLength(1)]
        [MaxLength(30)]
        public string TextValue { get; set; } = "Sample Text";
        public DateTime DateValue { get; set; } = DateTime.Now;
        public double NumericValue { get; set; } = 0.0;

        public ExcelNumberFormat NumberFormat { get; set; } = ExcelNumberFormat.Number;
        public ExcelNumberFormat DateFormat { get; set; } = ExcelNumberFormat.ShortDate;

        internal readonly List<object> NumericFormats = new()
        {
            new { Text = "General", Value = ExcelNumberFormat.General },
            new { Text = "Number (with 2 decimal places)", Value = ExcelNumberFormat.Number },
            new { Text = "Currency ($)", Value = ExcelNumberFormat.Currency },
            new { Text = "Accounting", Value = ExcelNumberFormat.Accounting },
            new { Text = "Percentage", Value = ExcelNumberFormat.Percentage },
            new { Text = "Scientific", Value = ExcelNumberFormat.Scientific }
        };

        internal readonly List<object> DateFormats = new()
        {
            new { Text = "Short Date", Value = ExcelNumberFormat.ShortDate },
            new { Text = "Long Date", Value = ExcelNumberFormat.LongDate }
        };
    }
}
