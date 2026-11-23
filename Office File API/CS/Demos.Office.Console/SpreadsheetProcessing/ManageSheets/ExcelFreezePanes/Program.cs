using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Freeze_Panes.xlsx";
    const int RowOffset = 1;
    const int ColumnOffset = 1;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FreezePanes(inputStream, RowOffset, ColumnOffset, OutputFormat);
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

    static Stream FreezePanes(Stream inputStream, int rowOffset, int columnOffset, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets.First();
        if(rowOffset > 0 && columnOffset > 0)
            worksheet.FreezePanes(rowOffset - 1, columnOffset - 1);
        else if(rowOffset == 0)
            worksheet.FreezeColumns(columnOffset - 1);
        else
            worksheet.FreezeRows(rowOffset - 1);
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
