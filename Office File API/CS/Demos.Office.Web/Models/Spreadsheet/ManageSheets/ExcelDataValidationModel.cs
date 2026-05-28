namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelDataValidationModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Data_Validation";

        public bool ShowInputMessage { get; set; } = true;
        public bool ShowErrorMessage { get; set; } = true;

        public IEnumerable<DataValidationRule> SelectedRules { get; set; } = new List<DataValidationRule>() {
            DataValidationRule.DecimalNumberRange,
            DataValidationRule.Formula,
            DataValidationRule.TextLength,
            DataValidationRule.ListCreatedInCode,
            DataValidationRule.ExistingWorksheetValues,
        };

        internal readonly List<object> DataValidationRules = new()
        {
            new { Text = "Restrict data entry in the \"F4:F11\" cell range to a decimal number from 10 to 40", Value = DataValidationRule.DecimalNumberRange },
            new { Text = "Use a formula to restrict data entry in the \"B4:B11\" cell range to a 5-digit number", Value = DataValidationRule.Formula },
            new { Text = "Restrict data entry in the \"D4:D11\" cell range to 3 symbols", Value = DataValidationRule.TextLength },
            new { Text = "Restrict data entry in the \"A4:A11\" cell range to \"PASS\" and \"FAIL\" values", Value = DataValidationRule.ListCreatedInCode },
            new { Text = "Restrict data entry in the \"E4:E11\" cell range to existing worksheet values", Value = DataValidationRule.ExistingWorksheetValues },
        };
    }
    public enum DataValidationRule {
        DecimalNumberRange,
        Formula,
        TextLength,
        ListCreatedInCode,
        ExistingWorksheetValues,
    }
}
