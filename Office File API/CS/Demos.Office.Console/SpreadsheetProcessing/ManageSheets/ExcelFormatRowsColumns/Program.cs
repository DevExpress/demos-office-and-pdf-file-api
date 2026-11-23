using DevExpress.Office;
using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.xlsx";
    const bool FormatRows = true;
    const FormatOperationType OperationType = FormatOperationType.Add;
    const int RowCount = 1;
    const int ColumnCount = 1;
    const int StartRowIndex = 1;
    const int StartColumnIndex = 1;
    const double RowHeight = 1;
    const double ColumnWidth = 1;
    const string DocumentUnitName = "Inch";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Process(inputStream);
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

    static Stream Process(Stream inputStream) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets.First();
        if(FormatRows) {
            int start = StartRowIndex - 1;
            switch(OperationType) {
                case FormatOperationType.Add:
                    worksheet.Rows.Insert(start, RowCount);
                    break;
                case FormatOperationType.Remove:
                    worksheet.Rows.Remove(start, RowCount);
                    break;
                case FormatOperationType.SetSize:
                    workbook.Unit = GetDocumentUnit(DocumentUnitName);
                    for(int i = start; i < start + RowCount; i++)
                        worksheet.Rows[i].Height = RowHeight;
                    break;
                case FormatOperationType.Hide:
                    worksheet.Rows.Hide(start, start + RowCount - 1);
                    break;
                case FormatOperationType.AutoFit:
                    worksheet.Rows.AutoFit(start, start + RowCount - 1);
                    break;
            }
        }
        else {
            int start = StartColumnIndex - 1;
            switch(OperationType) {
                case FormatOperationType.Add:
                    worksheet.Columns.Insert(start, ColumnCount);
                    break;
                case FormatOperationType.Remove:
                    worksheet.Columns.Remove(start, ColumnCount);
                    break;
                case FormatOperationType.SetSize:
                    workbook.Unit = GetDocumentUnit(DocumentUnitName);
                    for(int i = start; i < start + ColumnCount; i++)
                        worksheet.Columns[i].Width = ColumnWidth;
                    break;
                case FormatOperationType.Hide:
                    worksheet.Columns.Hide(start, start + ColumnCount - 1);
                    break;
                case FormatOperationType.AutoFit:
                    worksheet.Columns.AutoFit(start, start + ColumnCount - 1);
                    break;
            }
        }
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, OutputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static DocumentUnit GetDocumentUnit(string unit) => unit.ToLowerInvariant() switch {
        "centimeter" => DocumentUnit.Centimeter,
        "document" => DocumentUnit.Document,
        "inch" => DocumentUnit.Inch,
        "millimeter" => DocumentUnit.Millimeter,
        "point" => DocumentUnit.Point,
        _ => DocumentUnit.Document
    };
}

enum FormatOperationType {
    Add,
    Remove,
    SetSize,
    Hide,
    AutoFit
}
