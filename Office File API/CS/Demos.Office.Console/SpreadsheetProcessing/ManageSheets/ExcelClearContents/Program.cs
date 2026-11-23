using DevExpress.Spreadsheet;

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
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet firstWorksheet = workbook.Worksheets.First();
        CellRange rangeToClear = firstWorksheet.GetUsedRange();
        switch(clearMode) {
            case ClearMode.All:
                firstWorksheet.Clear(rangeToClear);
                break;
            case ClearMode.CellContent:
                firstWorksheet.ClearContents(rangeToClear);
                rangeToClear.Value = null;
                break;
            case ClearMode.CellFormatting:
                firstWorksheet.ClearFormats(rangeToClear);
                rangeToClear.Style = workbook.Styles.DefaultStyle;
                break;
            case ClearMode.CellHyperlinks:
                firstWorksheet.ClearHyperlinks(rangeToClear);
                foreach(Hyperlink hyperlink in firstWorksheet.Hyperlinks.GetHyperlinks(rangeToClear).ToArray())
                    firstWorksheet.Hyperlinks.Remove(hyperlink);
                break;
            case ClearMode.CellComments:
                firstWorksheet.ClearComments(rangeToClear);
                break;
        }
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
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
