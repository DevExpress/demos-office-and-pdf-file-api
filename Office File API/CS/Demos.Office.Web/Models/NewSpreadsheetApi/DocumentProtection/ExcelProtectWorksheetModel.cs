using System.ComponentModel.DataAnnotations;
using System.Text.RegularExpressions;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;
using DevExpress.Docs.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentProtection {
    public class ExcelProtectWorksheetModel : ExcelFileModelBase, IValidatableObject {
        [Required(ErrorMessage = "Password is required")]
        [MinLength(1, ErrorMessage = "Password cannot be empty")]
        public string WorksheetPassword { get; set; } = "123";

        public override string OutputName { get; set; } = "Protected_Document";

        public string WorksheetRange { get; set; } = "1-2";

        public IEnumerable<WorksheetProtectionPermissions> SelectedWorksheetProtectionPermissions { get; set; } = new List<WorksheetProtectionPermissions> {
            WorksheetProtectionPermissions.SelectLockedCells,
            WorksheetProtectionPermissions.SelectUnlockedCells
        };

        internal readonly List<EnumWrapper<WorksheetProtectionPermissions>> WorksheetPermissions = new() {
            new(WorksheetProtectionPermissions.AutoFilters, "Use Auto Filter"),
            new(WorksheetProtectionPermissions.DeleteColumns, "Delete columns"),
            new(WorksheetProtectionPermissions.DeleteRows, "Delete rows"),
            new(WorksheetProtectionPermissions.FormatCells, "Format cells"),
            new(WorksheetProtectionPermissions.FormatColumns, "Format columns"),
            new(WorksheetProtectionPermissions.FormatRows, "Format rows"),
            new(WorksheetProtectionPermissions.InsertColumns, "Insert columns"),
            new(WorksheetProtectionPermissions.InsertHyperlinks, "Insert hyperlinks"),
            new(WorksheetProtectionPermissions.InsertRows, "Insert rows"),
            new(WorksheetProtectionPermissions.Objects, "Edit objects"),
            new(WorksheetProtectionPermissions.PivotTables, "Use Pivot Table reports"),
            new(WorksheetProtectionPermissions.Scenarios, "Edit scenarios"),
            new(WorksheetProtectionPermissions.SelectLockedCells, "Select locked cells"),
            new(WorksheetProtectionPermissions.SelectUnlockedCells, "Select unlocked cells"),
            new(WorksheetProtectionPermissions.Sort, "Sort"),
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if(string.IsNullOrWhiteSpace(WorksheetRange)) {
                yield return new ValidationResult(
                    "WorksheetRange is required.",
                    new[] { nameof(WorksheetRange) });
                yield break;
            }

            const string pattern = @"^\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*$";
            if(!Regex.IsMatch(WorksheetRange, pattern)) {
                yield return new ValidationResult(
                    "WorksheetRange must be a comma-separated list of numbers or ranges (e.g. 1,3,5-7, 10-12).",
                    new[] { nameof(WorksheetRange) });
            }
        }
    }
}
