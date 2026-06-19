using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = ClearPages(
                File.OpenRead(inputPath),
                clearAnnotations: true,
                clearGraphics: true,
                clearImages: true,
                clearText: true,
                pageRange: []);
            SaveDocument(outputStream, Path.Combine(outputDirectory, "Document.pdf"));

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

    static Stream ClearPages(Stream documentStream, bool clearAnnotations, bool clearGraphics,
        bool clearImages, bool clearText, IEnumerable<int> pageRange) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, processor.Document.Pages.Count);

        foreach(var index in pageRange) {
            if(index < 0 || index >= processor.Document.Pages.Count)
                continue;

            var page = processor.Document.Pages[index];
            var options = new PdfClearContentOptions() {
                ClearAnnotations = clearAnnotations,
                ClearGraphics = clearGraphics,
                ClearImages = clearImages,
                ClearText = clearText
            };

            processor.DocumentFacade.Pages[index].ClearContent(page.MediaBox, options);
        }

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
