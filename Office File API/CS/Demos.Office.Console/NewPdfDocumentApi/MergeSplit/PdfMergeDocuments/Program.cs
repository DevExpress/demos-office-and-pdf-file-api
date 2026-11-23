using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string FirstInputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    static readonly string SecondInputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "MergedDocuments.pdf";

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var firstDocument = File.OpenRead(Path.Combine(AppContext.BaseDirectory, FirstInputFilePath));
            using var secondDocument = File.OpenRead(Path.Combine(AppContext.BaseDirectory, SecondInputFilePath));
            using var outputStream = MergeDocuments([firstDocument, secondDocument]);

            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static Stream MergeDocuments(IEnumerable<Stream> documentStreams) {
        using var doc = new PdfDocument();
        foreach(var stream in documentStreams)
            doc.AppendDocument(stream);
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
