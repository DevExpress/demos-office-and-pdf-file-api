using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Calculated_Document.pdf";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;
    static readonly CalculationMode? DocumentCalculationMode = null;
    static readonly bool ExportToPdf = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = CalculateWorkbook(inputStream, DocumentCalculationMode, ExportToPdf, OutputFormat);
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

    static Stream CalculateWorkbook(Stream inputStream, CalculationMode? documentCalculationMode, bool exportToPdf, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        workbook.Calculate();

        if(documentCalculationMode.HasValue)
            workbook.CalculationOptions.Mode = documentCalculationMode.Value;

        var outputStream = new MemoryStream();

        if(exportToPdf)
            workbook.ExportToPdf(outputStream);
        else
            workbook.Save(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);

        return outputStream;
    }
}
