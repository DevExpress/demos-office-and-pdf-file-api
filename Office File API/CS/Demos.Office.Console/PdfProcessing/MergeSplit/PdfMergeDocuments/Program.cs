using DevExpress.Pdf;

internal static class Program {
    static readonly string[] InputFilePaths = {
        Path.Combine("Documents", "PDF", "Sample_Main.pdf"),
        Path.Combine("Documents", "PDF", "Sample_Alternative.pdf")
    };
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Document.pdf";

    static int Main() {
        try {
            var inputPaths = InputFilePaths
                .Select(path => Path.Combine(AppContext.BaseDirectory, path))
                .ToArray();
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            var documentStreams = inputPaths
                .Select(File.OpenRead)
                .Cast<Stream>()
                .ToArray();

            try {
                using var outputStream = MergeDocuments(documentStreams);
                using var fileStream = File.Create(outputPath);
                outputStream.CopyTo(fileStream);
            }
            finally {
                foreach(var stream in documentStreams)
                    stream.Dispose();
            }

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static MemoryStream MergeDocuments(IEnumerable<Stream> documentStreams) {
        using var processor = new PdfDocumentProcessor();

        processor.CreateEmptyDocument();

        foreach(var stream in documentStreams)
            processor.AppendDocument(stream);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;

        return outputStream;
    }
}
