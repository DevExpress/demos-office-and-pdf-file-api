using DevExpress.Spreadsheet;

class Program {
    static readonly string[] InputFilePaths = {
        Path.Combine("Documents", "InvestmentPortfolio.xlsx"),
        Path.Combine("Documents", "DocumentForProtection.xlsx")
    };
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Merged_Document.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPaths = InputFilePaths.Select(path => Path.Combine(AppContext.BaseDirectory, path)).ToList();
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            var inputStreams = inputPaths.Select(File.OpenRead).Cast<Stream>().ToList();
            try {
                using var outputStream = MergeWorkbooks(inputStreams, OutputFormat);
                using var fileStream = File.Create(outputPath);
                outputStream.CopyTo(fileStream);
            }
            finally {
                foreach(var inputStream in inputStreams)
                    inputStream.Dispose();
            }

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream MergeWorkbooks(List<Stream> inputStreams, DocumentFormat outputFormat) {
        using var finalWorkbook = new Workbook();

        finalWorkbook.LoadDocument(inputStreams[0]);

        for(var index = 1; index < inputStreams.Count; index++) {
            using var currentWorkbook = new Workbook();
            currentWorkbook.LoadDocument(inputStreams[index]);
            finalWorkbook.Append(currentWorkbook);
        }

        var outputStream = new MemoryStream();
        finalWorkbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
