using DevExpress.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.Formatting {
    public class ExcelConditionalFormattingModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "Conditional_Formatting";

        public ConditionalFormattingRuleType RuleType { get; set; } = ConditionalFormattingRuleType.Average;
        public ConditionalFormattingAverageCondition AverageCondition { get; set; } = ConditionalFormattingAverageCondition.Above;
        public ConditionalFormattingDataBarAxisPosition DataBarAxisPosition { get; set; } = ConditionalFormattingDataBarAxisPosition.Middle;

        public string AverageRuleBackColorHex { get; set; }
        internal object AverageRuleBackColor { get; set; } = "#FAF7AA";
        public string AverageRuleForeColorHex { get; set; }
        internal object AverageRuleForeColor { get; set; } = "#FF0000";

        public string ColorScale2RuleMinPointColorHex { get; set; }
        internal object ColorScale2RuleMinPointColor { get; set; } = "#9DE9FA";
        public string ColorScale2RuleMaxPointColorHex { get; set; }
        internal object ColorScale2RuleMaxPointColor { get; set; } = "#FFF6A9";

        public string DataBarRulePositiveColorHex { get; set; }
        internal object DataBarRulePositiveColor { get; set; } = "#00FF00";
        public string DataBarRuleNegativeColorHex { get; set; }
        internal object DataBarRuleNegativeColor { get; set; } = "#FF0000";

        public string FormulaRuleRowBackColorHex { get; set; }
        internal object FormulaRuleRowBackColor { get; set; } = "#BCDAF7";

        internal readonly List<object> ConditionalFormattingAverageConditions = new(){
            new { Text = "Above", Value = ConditionalFormattingAverageCondition.Above },
            new { Text = "Above Or Equal", Value = ConditionalFormattingAverageCondition.AboveOrEqual },
            new { Text = "Below", Value = ConditionalFormattingAverageCondition.Below },
            new { Text = "Below Or Equal", Value = ConditionalFormattingAverageCondition.BelowOrEqual }
        };
        internal readonly List<object> ConditionalFormattingDataBarAxisPositions = new(){
            new { Text = "Automatic", Value = ConditionalFormattingDataBarAxisPosition.Automatic },
            new { Text = "Middle", Value = ConditionalFormattingDataBarAxisPosition.Middle },
            new { Text = "None", Value = ConditionalFormattingDataBarAxisPosition.None },
        };
        internal readonly List<object> ConditionalFormattingRuleTypes = new(){
            new { Text = "Above/Below Average", Value = ConditionalFormattingRuleType.Average },
            new { Text = "Two-Color Scale", Value = ConditionalFormattingRuleType.ColorScale2 },
            new { Text = "Data Bars", Value = ConditionalFormattingRuleType.DataBar },
            new { Text = "Formula", Value = ConditionalFormattingRuleType.Formula },
        };
    }
    public enum ConditionalFormattingRuleType {
        Average,
        ColorScale2,
        DataBar,
        Formula
    }
}
