using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    static readonly int[] PageRange = Array.Empty<int>();

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var documentStream = File.OpenRead(inputPath);
            var documents = SplitDocuments(documentStream, PageRange);
            for(int index = 0; index < documents.Count; index++) {
                using var stream = documents[index];
                SaveDocument(stream, Path.Combine(outputDirectory, $"document_{index + 1}.pdf"));
            }

            Console.WriteLine($"Created {outputDirectory}");
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

    static IReadOnlyList<Stream> SplitDocuments(Stream documentStream, IEnumerable<int> pageRange) {
        using var source = new PdfDocument(documentStream);

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, source.Pages.Count);

        var result = new List<Stream>();
        foreach(var index in pageRange) {
            if(index < 0 || index >= source.Pages.Count)
                continue;

            using var pageDocument = new PdfDocument();
            pageDocument.Pages.Add(source.Pages[index].Clone());

            var outputStream = new MemoryStream();
            pageDocument.Save(outputStream);
            outputStream.Position = 0;
            result.Add(outputStream);
        }

        return result;
    }
}
