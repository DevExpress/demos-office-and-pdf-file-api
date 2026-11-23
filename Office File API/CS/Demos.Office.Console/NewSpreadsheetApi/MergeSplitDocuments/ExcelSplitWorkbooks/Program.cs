using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "TrendlineAnalysis_Template.xlsx");
    const string OutputDirectoryName = "Output";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;
    const bool SplitVisibleOnly = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);

            if(Directory.Exists(outputDirectory))
                Directory.Delete(outputDirectory, recursive: true);
            Directory.CreateDirectory(outputDirectory);

            using var inputStream = File.OpenRead(inputPath);
            var documentStreams = SplitWorkbook(SplitVisibleOnly, inputStream, OutputFormat);

            if(documentStreams.Count == 0)
                throw new InvalidOperationException("Failed to split workbook.");

            for(var index = 0; index < documentStreams.Count; index++) {
                var outputPath = Path.Combine(outputDirectory, $"worksheet_{index + 1}.xlsx");
                using var fileStream = File.Create(outputPath);
                using var outputStream = documentStreams[index];
                outputStream.CopyTo(fileStream);
            }

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static IReadOnlyList<Stream> SplitWorkbook(bool splitVisibleOnly, Stream inputStream, DocumentFormat outputFormat) {
        using var sourceWorkbook = new Workbook(inputStream);

        var options = new WorkbookExtractionOptions { UnhideSheets = true };

        var documentStreams = new List<Stream>();
        for(var index = 0; index < sourceWorkbook.Sheets.Count; index++) {
            if(!ShouldExtractSheet(splitVisibleOnly, sourceWorkbook.Sheets[index]))
                continue;

            using var targetWorkbook = sourceWorkbook.Extract(options, index);

            targetWorkbook.Calculate();

            var outputStream = new MemoryStream();

            targetWorkbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            documentStreams.Add(outputStream);
        }
        return documentStreams.AsReadOnly();
    }

    static bool ShouldExtractSheet(bool splitVisibleOnly, Sheet sheet) {
        if(sheet.Type == SheetType.ChartSheet)
            return false;
        if(splitVisibleOnly)
            return sheet.Visibility == SheetVisibility.Visible;
        return sheet.Visibility != SheetVisibility.VeryHidden;
    }
}
