using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelWorksheetOperationsModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Modified_Document";

        public OperationType OperationType { get; set; } = OperationType.Add;
        [MinLength(1)]
        [MaxLength(32)]
        public string WorksheetName { get; set; } = "New Worksheet";

        public string WorksheetRange { get; set; } = string.Empty;

        public bool WorksheetRangeEnabled => OperationType != OperationType.Add;

        // Conditional validation for WorksheetRange
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if (WorksheetRangeEnabled) {
                if (string.IsNullOrWhiteSpace(WorksheetRange)) {
                    yield return new ValidationResult(
                        "WorksheetRange is required when the selected Operation Type is 'Remove Worksheet' or 'Duplicate Worksheet'.",
                        new[] { nameof(WorksheetRange) });
                    yield break;
                }

                // Accept patterns like: 1,2,4-6, 10-12,15
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
