using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Table_Document.xlsx";
    const int NumberOfRows = 5;
    const int NumberOfColumns = 5;
    const bool CreateCustomTableStyle = false;
    const bool ShowTotals = true;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = GenerateDocumentWithTable(NumberOfRows, NumberOfColumns, true, CreateCustomTableStyle, ShowTotals, OutputFormat);
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

    static Stream GenerateDocumentWithTable(int rowCount, int columnCount, bool hasHeaders, bool useCustomTableStyle, bool showTotals, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets.First();
        var range = worksheet.Range.FromLTRB(0, 0, columnCount - 1, hasHeaders ? rowCount : rowCount - 1);

        if(hasHeaders) {
            for(int column = 0; column < columnCount; column++)
                range[0, column].SetValue($"Header {column + 1}");
        }

        int startRow = hasHeaders ? 1 : 0;
        for(int row = 0; row < rowCount; row++) {
            for(int column = 0; column < columnCount; column++) {
                var cell = range[startRow + row, column];
                cell.SetValue(column == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{column + 1}");
            }
        }

        var table = worksheet.Tables.Add(range, hasHeaders);
        var tableStyles = workbook.TableStyles;
        table.Style = useCustomTableStyle ? CreateCustomStyle(tableStyles, "CustomTableStyle") : tableStyles[BuiltInTableStyleId.TableStyleDark9];

        if(showTotals) {
            table.ShowTotals = true;
            table.Columns.Last().TotalRowFunction = TotalRowFunction.Count;
        }

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static TableStyle CreateCustomStyle(TableStyleCollection tableStyles, string styleName) {
        var customTableStyle = tableStyles.Add(styleName);
        customTableStyle.BeginUpdate();
        try {
            customTableStyle.TableStyleElements[TableStyleElementType.WholeTable].Font.Color = Color.FromArgb(107, 107, 107);

            var headerRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.HeaderRow];
            headerRowStyle.Fill.BackgroundColor = Color.FromArgb(64, 66, 166);
            headerRowStyle.Font.Color = Color.White;
            headerRowStyle.Font.Bold = true;

            var totalRowStyle = customTableStyle.TableStyleElements[TableStyleElementType.TotalRow];
            totalRowStyle.Fill.BackgroundColor = Color.FromArgb(115, 193, 211);
            totalRowStyle.Font.Color = Color.White;
            totalRowStyle.Font.Bold = true;

            var secondRowStripeStyle = customTableStyle.TableStyleElements[TableStyleElementType.SecondRowStripe];
            secondRowStripeStyle.Fill.BackgroundColor = Color.FromArgb(234, 234, 234);
            secondRowStripeStyle.StripeSize = 1;
        }
        finally {
            customTableStyle.EndUpdate();
        }

        return customTableStyle;
    }
}
