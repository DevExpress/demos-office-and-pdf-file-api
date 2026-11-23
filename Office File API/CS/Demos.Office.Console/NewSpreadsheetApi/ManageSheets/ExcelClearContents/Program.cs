using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.xlsx";
    const ClearMode ContentClearMode = ClearMode.All;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ClearContents(inputStream, ContentClearMode, OutputFormat);
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

    static Stream ClearContents(Stream inputStream, ClearMode clearMode, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet firstWorksheet = workbook.Worksheets[0];

        CellRange rangeToClear = firstWorksheet.GetUsedRange();

        ClearContentType contentToClear = clearMode switch {
            ClearMode.CellContent => ClearContentType.Values,
            ClearMode.CellFormatting => ClearContentType.Formats,
            ClearMode.CellHyperlinks => ClearContentType.Hyperlinks,
            ClearMode.CellComments => ClearContentType.Comments,
            _ => ClearContentType.All
        };

        rangeToClear.Clear(contentToClear);

        var outputStream = new MemoryStream();
        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}

enum ClearMode {
    All,
    CellContent,
    CellFormatting,
    CellHyperlinks,
    CellComments
}
