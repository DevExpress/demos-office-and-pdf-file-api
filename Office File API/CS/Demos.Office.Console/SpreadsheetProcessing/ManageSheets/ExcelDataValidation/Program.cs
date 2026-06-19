using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "DataValidation_Template.xlsx");
    const string OutputFileName = "Data_Validation.xlsx";
    static readonly bool ShowInputMessage = true;
    static readonly bool ShowErrorMessage = true;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ApplyValidation(inputStream, OutputFormat);
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

    static Stream ApplyValidation(Stream inputStream, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets.First();

        var decimalRule = worksheet.DataValidations.Add(worksheet["F4:F11"], DataValidationType.Decimal, DataValidationOperator.Between, 10, 40);
        Configure(decimalRule, "Decimal Number", "Enter decimal numbers between 10 and 40", "Wrong Number", "The value you entered is not valid. Enter a decimal number between 10 and 40.");

        var formulaRule = worksheet.DataValidations.Add(worksheet["B4:B11"], DataValidationType.Custom, "=AND(ISNUMBER(B4),LEN(B4)=5)");
        Configure(formulaRule, "Employee Id", "Enter 5-digit number", "Wrong Employee Id", "The value you entered is not valid. Use 5-digit number for the employee ID.");

        var textLengthRule = worksheet.DataValidations.Add(worksheet["D4:D11"], DataValidationType.TextLength, DataValidationOperator.Equal, 3);
        Configure(textLengthRule, "Alias", "Enter a text value that contains 3 symbols", "Wrong Alias", "The value you entered is not valid. Entered value must contain 3 symbols.");

        var listRule = worksheet.DataValidations.Add(worksheet["A4:A11"], DataValidationType.List, "PASS, FAIL");
        Configure(listRule, "Assessment", "Enter a \"PASS\" or \"FAIL\" value", "Wrong Value", "The value you entered is not valid. Valid values: \"PASS\", \"FAIL\".");

        var existingValuesRule = worksheet.DataValidations.Add(worksheet["E4:E11"], DataValidationType.List, ValueObject.FromRange(worksheet["H4:H9"].GetRangeWithAbsoluteReference()));
        Configure(existingValuesRule, "Department", "Select a Department from the drop-down list", "Wrong Department", "The value you entered is not valid. Select a valid Department value from the drop-down list.");

        HighlightRanges(worksheet);
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void Configure(DataValidation validation, string inputTitle, string inputMessage, string errorTitle, string errorMessage) {
        if(ShowInputMessage) {
            validation.ShowInputMessage = true;
            validation.InputTitle = inputTitle;
            validation.InputMessage = inputMessage;
        }
        if(ShowErrorMessage) {
            validation.ShowErrorMessage = true;
            validation.ErrorStyle = DataValidationErrorStyle.Information;
            validation.ErrorTitle = errorTitle;
            validation.ErrorMessage = errorMessage;
        }
    }

    static void HighlightRanges(Worksheet worksheet) {
        worksheet["H4:H9"].FillColor = Color.LightGray;
        int[] colors = [0xFFC4C4, 0xFFD9D9, 0xFFF6F6, 0xFFECEC, 0xE9D3D3, 0xFFDFC4, 0xFFDAE9];
        for(int i = 0; i < worksheet.DataValidations.Count; i++)
            worksheet.DataValidations[i].Range.FillColor = Color.FromArgb(colors[i]);
    }
}
