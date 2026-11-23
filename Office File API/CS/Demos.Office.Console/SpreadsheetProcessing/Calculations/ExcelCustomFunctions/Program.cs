using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Functions;
using System.Globalization;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Custom_Functions.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var outputStream = CreateWorkbookWithCustomFunction(OutputFormat);
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

    static Stream CreateWorkbookWithCustomFunction(DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        SphereMassFunction customFunction = new SphereMassFunction();
        var globalFunctions = workbook.Functions.GlobalCustomFunctions;
        if(!globalFunctions.Contains(customFunction.Name))
            globalFunctions.Add(customFunction);

        workbook.BeginUpdate();
        try {
            Worksheet worksheet = workbook.Worksheets.First();
            worksheet.Name = "Custom User-Defined Functions";
            worksheet.Range["A1:H1"].ColumnWidthInCharacters = 12;
            worksheet.Range["A1:H1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

            worksheet.DefinedNames.Add("seawater", "1025");

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

            worksheet.Calculate();

            worksheet["D2"].NumberFormat = "#.##";
            worksheet["D2"].Value = worksheet["C2"].Value;

            worksheet["D3"].NumberFormat = "#.##";
            worksheet["D3"].Value = worksheet["C3"].Value;

            worksheet.Columns.AutoFit("B", "D");
        } finally {
            workbook.EndUpdate();
        }

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    public class SphereMassFunction : ICustomFunction {
        const string functionName = "SPHEREMASS";
        readonly ParameterInfo[] functionParameters;

        public SphereMassFunction() {
            functionParameters = [
                new ParameterInfo(ParameterType.Value, ParameterAttributes.Required),
                new ParameterInfo(ParameterType.Value, ParameterAttributes.Optional)
            ];
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
}
