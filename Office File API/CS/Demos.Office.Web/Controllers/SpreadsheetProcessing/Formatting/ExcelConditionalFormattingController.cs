#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Formatting;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Formatting {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelConditionalFormattingController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("conditional-formatting/format")]
        public async Task<IActionResult> Format(IFormFile file, [FromForm] ExcelConditionalFormattingModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Add Conditional Formatting Rule to Workbook
            Stream outputStream = new MemoryStream();
            switch(model.RuleType) {
                case ConditionalFormattingRuleType.Average:
                    outputStream = AddAverageConditionalFormatting(inputStream, ColorTranslator.FromHtml(model.AverageRuleBackColorHex ?? string.Empty), ColorTranslator.FromHtml(model.AverageRuleForeColorHex ?? string.Empty), model.AverageCondition, targetFormat);
                    break;
                case ConditionalFormattingRuleType.ColorScale2:
                    outputStream = AddColorScale2ConditionalFormatting(inputStream, ColorTranslator.FromHtml(model.ColorScale2RuleMinPointColorHex ?? string.Empty), ColorTranslator.FromHtml(model.ColorScale2RuleMaxPointColorHex ?? string.Empty), targetFormat);
                    break;
                case ConditionalFormattingRuleType.DataBar:
                    outputStream = AddDataBarConditionalFormatting(inputStream, ColorTranslator.FromHtml(model.DataBarRulePositiveColorHex ?? string.Empty), ColorTranslator.FromHtml(model.DataBarRuleNegativeColorHex ?? string.Empty), model.DataBarAxisPosition, targetFormat);
                    break;
                case ConditionalFormattingRuleType.Formula:
                    outputStream = AddFormulaExpressionConditionalFormatting(inputStream, ColorTranslator.FromHtml(model.FormulaRuleRowBackColorHex ?? string.Empty), targetFormat);
                    break;
            }

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddAverageConditionalFormatting(Stream inputStream, Color averageRuleBackColor, Color averageRuleForeColor,
            ConditionalFormattingAverageCondition averageCondition, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            workbook.Calculate();
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                // Access the conditional formatting rules collection
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Create the rule highlighting values that are above the average in cells C2 through D15
                AverageConditionalFormatting averageRule = conditionalFormattings.AddAverageConditionalFormatting(worksheet.Range["$C$2:$D$15"], averageCondition);
                // Set the background color
                averageRule.Formatting.Fill.BackgroundColor = averageRuleBackColor;
                // Set the text color
                averageRule.Formatting.Font.Color = averageRuleForeColor;
                // Add an explanation to the created rule
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = $"Determine cost values that are above/below the average in the first and second quarters.";
            } finally {
                workbook.EndUpdate();
            }

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddColorScale2ConditionalFormatting(Stream inputStream, Color minPointColor, Color maxPointColor, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            workbook.Calculate();
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                // Access the conditional formatting rules collection
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the minimum threshold to the lowest value in the range of cells
                ConditionalFormattingValue minPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Set the maximum threshold to the highest value in the range of cells
                ConditionalFormattingValue maxPoint = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Create the two-color scale rule to differentiate low and high values in cells C2 through D15.
                // MinPointColor represents the lower values and MaxPointColor represents the higher values
                ColorScale2ConditionalFormatting colorScale2Rule = conditionalFormattings.AddColorScale2ConditionalFormatting(
                    worksheet.Range["$C$2:$D$15"], minPoint, minPointColor, maxPoint, maxPointColor);
                // Add an explanation to the created rule
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = $"Examine cost distribution using a gradation of two colors.";
            } finally {
                workbook.EndUpdate();
            }

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddDataBarConditionalFormatting(Stream inputStream, Color positiveColor, Color negativeColor,
            ConditionalFormattingDataBarAxisPosition dataBarAxisPosition, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            workbook.Calculate();
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                // Access the conditional formatting rules collection
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Set the value corresponding to the shortest bar to the lowest value
                ConditionalFormattingValue lowBound1 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Set the value corresponding to the longest bar to the highest value
                ConditionalFormattingValue highBound1 = conditionalFormattings.CreateValue(ConditionalFormattingValueType.MinMax);
                // Create the rule to compare values in cells E2 through E15 using data bars
                DataBarConditionalFormatting dataBarRule = conditionalFormattings.AddDataBarConditionalFormatting(
                    worksheet.Range["$E$2:$E$15"], lowBound1, highBound1, positiveColor);
                // Set the positive bar color
                dataBarRule.Color = positiveColor;
                // Set the positive bar border color
                dataBarRule.BorderColor = positiveColor;
                // Set the negative bar color
                dataBarRule.NegativeBarColor = negativeColor;
                // Set the negative bar border color
                dataBarRule.NegativeBarBorderColor = negativeColor;
                // Set the axis position
                dataBarRule.AxisPosition = dataBarAxisPosition;
                // Set the axis color to dark blue
                dataBarRule.AxisColor = Color.DarkBlue;
                // Add an explanation to the created rule
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Compare values in the \"Cost Trend\" column using data bars.";
            } finally {
                workbook.EndUpdate();
            }

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream AddFormulaExpressionConditionalFormatting(Stream inputStream, Color rowBackgroundColor, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            workbook.Calculate();
            workbook.BeginUpdate();
            try {
                Worksheet worksheet = workbook.Worksheets["cfBooks"];
                // Access the conditional formatting rules collection
                ConditionalFormattingCollection conditionalFormattings = worksheet.ConditionalFormattings;
                // Create the rule to shade alternate rows without applying a new style
                FormulaExpressionConditionalFormatting cfRule = conditionalFormattings.AddFormulaExpressionConditionalFormatting(
                    worksheet.Range["$A$2:$G$15"], "=MOD(ROW(),2)=1");
                // Set the background color
                cfRule.Formatting.Fill.BackgroundColor = rowBackgroundColor;
                // Add an explanation to the created rule
                CellRange ruleExplanation = worksheet.Range["A17:G18"];
                ruleExplanation.Value = "Shade alternate rows without applying a new style.";
            }
            finally {
                workbook.EndUpdate();
            }

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
