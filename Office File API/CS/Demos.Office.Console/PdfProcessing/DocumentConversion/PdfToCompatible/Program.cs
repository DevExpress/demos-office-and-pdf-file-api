using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document.pdf";
    const PdfCompatibility CompatibilityLevel = PdfCompatibility.PdfA2b;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = ConvertToCompatible(documentStream, CompatibilityLevel);
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

    static MemoryStream ConvertToCompatible(Stream documentStream, PdfCompatibility compatibilityLevel) {
        using var converter = new PdfDocumentConverter(documentStream);

        var outputStream = new MemoryStream();

        converter.Convert(compatibilityLevel);
        converter.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
