using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets {
    public class ExcelWorksheetOperationsModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Modified_Document";

        public OperationType OperationType { get; set; } = OperationType.Add;
        [MinLength(1)]
        [MaxLength(32)]
        public string WorksheetName { get; set; } = "New Worksheet";

        public string WorksheetRange { get; set; } = string.Empty;

        public bool WorksheetRangeEnabled => OperationType != OperationType.Add;

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if (WorksheetRangeEnabled) {
                if (string.IsNullOrWhiteSpace(WorksheetRange)) {
                    yield return new ValidationResult(
                        "WorksheetRange is required when the selected Operation Type is 'Remove Worksheet' or 'Duplicate Worksheet'.",
                        new[] { nameof(WorksheetRange) });
                    yield break;
                }

                const string pattern = @"^\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*$";
                if (!Regex.IsMatch(WorksheetRange, pattern)) {
                    yield return new ValidationResult(
                        "WorksheetRange must be a comma-separated list of numbers or ranges (e.g. 1,3,5-7, 10-12).",
                        new[] { nameof(WorksheetRange) });
                }
            }
        }

        internal readonly List<object> OperationTypes = new()
        {
            new { Text = "Add Worksheet", Value = OperationType.Add },
            new { Text = "Remove Worksheet", Value = OperationType.Remove },
            new { Text = "Duplicate Worksheet", Value = OperationType.Duplicate },
        };
    }
    public enum OperationType {
        Add,
        Remove,
        Duplicate
    }
}
