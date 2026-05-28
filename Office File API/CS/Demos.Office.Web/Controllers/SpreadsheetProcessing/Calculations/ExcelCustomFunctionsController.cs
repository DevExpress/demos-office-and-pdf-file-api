#region Code
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;
using System.Globalization;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.Calculations;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.Calculations {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelCustomFunctionsController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("custom-functions/create")]
        public async Task<IActionResult> CreateWorkbook([FromForm] ExcelCustomFunctionsModel model) {
            // Create input stream
            using var inputStream = new MemoryStream();

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate a workbook with a custom function
            var convertedStream = CreateWorkbookWithCustomFunction(targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        public Stream CreateWorkbookWithCustomFunction(DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();

            // Create a custom function and add it to the global Workbook scope
            SphereMassFunction customFunction = new SphereMassFunction();
            var globalFunctions = workbook.Functions.GlobalCustomFunctions;
            if(!globalFunctions.Contains(customFunction.Name))
                globalFunctions.Add(customFunction);

            workbook.BeginUpdate();
            try {
                // Use the first worksheet
                Worksheet worksheet = workbook.Worksheets.First();
                worksheet.Name = "Custom User-Defined Functions";
                // Set cell widths and alignment
                worksheet.Range["A1:H1"].ColumnWidthInCharacters = 12;
                worksheet.Range["A1:H1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                // Define a named range for seawater density
                worksheet.DefinedNames.Add("seawater", "1025");

                // Initialize worksheet data
                worksheet["A1"].Value = "Radius, m";
                worksheet["B1"].Value = "Material";
                worksheet["C1"].Value = "Mass, kg";
                worksheet["D1"].Value = "Function Result";

                worksheet["A2"].Value = 0.1;
                worksheet["B2"].Value = "";
                worksheet["C2"].FormulaInvariant = "=SPHEREMASS(A2)";
                worksheet["C2"].NumberFormat = "#.##";

                worksheet["A3"].Value = 0.1;
                worksheet["B3"].Value = "Seawater";
                worksheet["C3"].FormulaInvariant = "=SPHEREMASS(A3, seawater)";
                worksheet["C3"].NumberFormat = "#.##";

                worksheet["A5"].Value = "Note: Excel can't calculate custom user-defined functions. Cells \"C2\" and \"C3\" display the '#NAME?' after recalculation attempts.";
                worksheet["A6"].Value = "The results of the custom SPHEREMASS function calculation are displayed in cells \"D2\" and \"D3\".";

                // Calculate worksheet to evaluate formulas
                worksheet.Calculate();

                // Display the results of the custom function calculation in the "D2" and "D3" cells
                worksheet["D2"].NumberFormat = "#.##";
                worksheet["D2"].Value = worksheet["C2"].Value;

                worksheet["D3"].NumberFormat = "#.##";
                worksheet["D3"].Value = worksheet["C3"].Value;

                // Auto-fit columns
                worksheet.Columns.AutoFit("B", "D");
            } finally {
                workbook.EndUpdate();
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }

        // Custom SPHEREMASS function implementation
        public class SphereMassFunction : ICustomFunction {
            const string functionName = "SPHEREMASS";
            readonly ParameterInfo[] functionParameters;

            public SphereMassFunction() {
                // Missing optional parameters do not result in an error message
                this.functionParameters =
                    new ParameterInfo[] {
                        new ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
                        new ParameterInfo(ParameterType.Value, ParameterAttributes.Optional)
                    };
            }
            public string Name { get { return functionName; } }
            ParameterInfo[] IFunction.Parameters { get { return functionParameters; } }
            ParameterType IFunction.ReturnType { get { return ParameterType.Value; } }
            bool IFunction.Volatile { get { return false; } }

            ParameterValue IFunction.Evaluate(IList<ParameterValue> parameters, EvaluationContext context) {
                double radius;
                double density = 1000;
                ParameterValue radiusParameter;
                ParameterValue densityParameter;

                if(parameters.Count == 2) {
                    densityParameter = parameters[1];
                    if(densityParameter.IsError)
                        return densityParameter;
                    else
                        density = densityParameter.NumericValue;
                }

                radiusParameter = parameters[0];
                if(radiusParameter.IsError)
                    return radiusParameter;
                else
                    radius = radiusParameter.NumericValue;

                return (4 * Math.PI) / 3 * Math.Pow(radius, 3) * density;
            }
            string IFunction.GetName(CultureInfo culture) {
                return functionName;
            }
        }
        #endregion Code
    }
}
