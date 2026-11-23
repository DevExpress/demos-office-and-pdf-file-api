using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document.pdf";
    const PdfACompatibility Compatibility = PdfACompatibility.PdfA2b;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = ConvertToPdfA(documentStream, Compatibility, out var result);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Conversion Status : {result.Status}");
            if(result.Issues.Any()) {
                Console.WriteLine("Found Issues :");
                foreach(var issue in result.Issues)
                    Console.WriteLine($"{issue.Severity} : {issue.Message}");
            }

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static MemoryStream ConvertToPdfA(Stream documentStream, PdfACompatibility compatibility, out ConversionResult result) {
        using var document = new PdfDocument(documentStream);

        var outputStream = new MemoryStream();

        result = document.ConvertToPdfA(compatibility);
        document.Save(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
