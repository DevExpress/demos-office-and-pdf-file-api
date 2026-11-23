using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "ConditionalFormattingSample.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Conditional_Formatting.xlsx";
    const ConditionalFormattingRuleType RuleType = ConditionalFormattingRuleType.Average;
    const ConditionalFormattingAverageCondition AverageCondition = ConditionalFormattingAverageCondition.Above;
    const ConditionalFormattingDataBarAxisPosition DataBarAxisPosition = ConditionalFormattingDataBarAxisPosition.Middle;
    static readonly Color AverageRuleBackColor = ColorTranslator.FromHtml("#FAF7AA");
    static readonly Color AverageRuleForeColor = ColorTranslator.FromHtml("#FF0000");
    static readonly Color ColorScale2RuleMinPointColor = ColorTranslator.FromHtml("#9DE9FA");
    static readonly Color ColorScale2RuleMaxPointColor = ColorTranslator.FromHtml("#FFF6A9");
    static readonly Color DataBarRulePositiveColor = ColorTranslator.FromHtml("#00FF00");
    static readonly Color DataBarRuleNegativeColor = ColorTranslator.FromHtml("#FF0000");
    static readonly Color FormulaRuleRowBackColor = ColorTranslator.FromHtml("#BCDAF7");
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = RuleType switch {
                ConditionalFormattingRuleType.Average => AddAverageConditionalFormatting(inputStream, AverageRuleBackColor, AverageRuleForeColor, AverageCondition, OutputFormat),
                ConditionalFormattingRuleType.ColorScale2 => AddColorScale2ConditionalFormatting(inputStream, ColorScale2RuleMinPointColor, ColorScale2RuleMaxPointColor, OutputFormat),
                ConditionalFormattingRuleType.DataBar => AddDataBarConditionalFormatting(inputStream, DataBarRulePositiveColor, DataBarRuleNegativeColor, DataBarAxisPosition, OutputFormat),
                ConditionalFormattingRuleType.Formula => AddFormulaExpressionConditionalFormatting(inputStream, FormulaRuleRowBackColor, OutputFormat),
                _ => AddAverageConditionalFormatting(inputStream, AverageRuleBackColor, AverageRuleForeColor, AverageCondition, OutputFormat)
            };
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream AddAverageConditionalFormatting(Stream inputStream, Color averageRuleBackColor, Color averageRuleForeColor,
        ConditionalFormattingAverageCondition averageCondition, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        workbook.Calculate();
        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
            AverageConditionalFormatting averageRule = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range["$C$2:$D$15"], averageCondition);
            averageRule.Formatting.Fill.BackgroundColor = averageRuleBackColor;
            averageRule.Formatting.Font.Color = averageRuleForeColor;
            worksheet.Range["A17:G18"].Value = "Determine cost values that are above/below the average in the first and second quarters.";
        }
        finally {
            workbook.EndUpdate();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddColorScale2ConditionalFormatting(Stream inputStream, Color minPointColor, Color maxPointColor, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        workbook.Calculate();
        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
            ConditionalFormattingValue minPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
            ConditionalFormattingValue maxPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
            conditionalFormattings.AddColorScale2ConditionalFormatting(worksheet.Range["$C$2:$D$15"], minPoint, minPointColor, maxPoint, maxPointColor);
            worksheet.Range["A17:G18"].Value = "Examine cost distribution using a gradation of two colors.";
        }
        finally {
            workbook.EndUpdate();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddDataBarConditionalFormatting(Stream inputStream, Color positiveColor, Color negativeColor,
        ConditionalFormattingDataBarAxisPosition dataBarAxisPosition, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        workbook.Calculate();
        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
            ConditionalFormattingValue lowBound = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
            ConditionalFormattingValue highBound = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
            DataBarConditionalFormatting dataBarRule = conditionalFormattings.AddDataBarConditionalFormatting(worksheet.Range["$E$2:$E$15"], lowBound, highBound, positiveColor);
            dataBarRule.Color = positiveColor;
            dataBarRule.BorderColor = positiveColor;
            dataBarRule.NegativeBarColor = negativeColor;
            dataBarRule.NegativeBarBorderColor = negativeColor;
            dataBarRule.AxisPosition = dataBarAxisPosition;
            dataBarRule.AxisColor = Color.DarkBlue;
            worksheet.Range["A17:G18"].Value = "Compare values in the \"Cost Trend\" column using data bars.";
        }
        finally {
            workbook.EndUpdate();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream AddFormulaExpressionConditionalFormatting(Stream inputStream, Color rowBackgroundColor, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        workbook.Calculate();
        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets["cfBooks"];
            ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
            FormulaExpressionConditionalFormatting rule = conditionalFormattings.AddFormulaExpressionConditionalFormatting(worksheet.Range["$A$2:$G$15"], "=MOD(ROW(),2)=1");
            rule.Formatting.Fill.BackgroundColor = rowBackgroundColor;
            worksheet.Range["A17:G18"].Value = "Shade alternate rows without applying a new style.";
        }
        finally {
            workbook.EndUpdate();
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }
}

enum ConditionalFormattingRuleType {
    Average,
    ColorScale2,
    DataBar,
    Formula
}
