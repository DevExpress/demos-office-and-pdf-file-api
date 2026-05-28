using System.ComponentModel.DataAnnotations;
using System.Text;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations {
    public class ExcelCircularReferencesModel : ExcelFileModelBase {
        public override string OutputFormat { get; set; } = "txt";
        public override string OutputName { get; set; } = "Circular_References";

        [Range(1, 10, ErrorMessage = "Max Iteration Count must be between 1 and 10.")]
        public int MaxIterationCount { get; set; } = 5;
    }
}
