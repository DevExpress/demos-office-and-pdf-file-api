using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets {
    public class ExcelFreezePanesModel : ExcelFileModelBase, IValidatableObject {
        public override string OutputName { get; set; } = "Freeze_Panes";
        [Range(0, 1048576)]
        public int RowOffset { get; set; } = 1;
        [Range(0, 16384)]
        public int ColumnOffset { get; set; } = 1;
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            if(RowOffset == 0 && ColumnOffset == 0) {
                yield return new ValidationResult(
                    "At least one parameter must be greater than 0.",
                    new[] { nameof(RowOffset), nameof(ColumnOffset) });
                yield break;
            }
        }
    }
}
