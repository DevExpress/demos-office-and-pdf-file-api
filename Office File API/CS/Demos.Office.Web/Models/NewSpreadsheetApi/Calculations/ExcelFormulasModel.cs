using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.Calculations {
    public class ExcelFormulasModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Formulas";

        [RegularExpression(@"^(ReferencesInFormulas|NamedRangesInFormulas|CreateNamedFormulas|FunctionsInFormulas|CreateSharedFormulas|CreateArrayFormulas)$",
        ErrorMessage = "Invalid formula type. Supported modes: ReferencesInFormulas, NamedRangesInFormulas, CreateNamedFormulas, FunctionsInFormulas, CreateSharedFormulas, CreateArrayFormulas")]
        public FormulaType FormulaType { get; set; } = FormulaType.ReferencesInFormulas;

        internal readonly List<object> FormulaTypes = new(){
            new { Text = "Formulas with Cell and Worksheet References", Value = FormulaType.ReferencesInFormulas },
            new { Text = "Formulas with Range Names", Value = FormulaType.NamedRangesInFormulas },
            new { Text = "Named Formulas", Value = FormulaType.CreateNamedFormulas },
            new { Text = "Formulas with Functions and Nested Functions", Value = FormulaType.FunctionsInFormulas },
            new { Text = "Shared Formulas", Value = FormulaType.CreateSharedFormulas },
            new { Text = "Array Formulas", Value = FormulaType.CreateArrayFormulas },
        };
    }

    public enum FormulaType {
        ReferencesInFormulas,
        NamedRangesInFormulas,
        CreateNamedFormulas,
        FunctionsInFormulas,
        CreateSharedFormulas,
        CreateArrayFormulas
    }
}
