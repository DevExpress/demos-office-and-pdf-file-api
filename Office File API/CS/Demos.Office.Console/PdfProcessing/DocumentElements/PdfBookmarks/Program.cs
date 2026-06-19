using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);

            using var addedBookmarksStream = AddBookmarks(new MemoryStream(inputBytes));
            var addedBookmarksBytes = ReadAllBytes(addedBookmarksStream);
            SaveDocument(new MemoryStream(addedBookmarksBytes), Path.Combine(outputDirectory, "bookmarks_added.pdf"));
            SaveDocument(DeleteBookmarks(new MemoryStream(addedBookmarksBytes)), Path.Combine(outputDirectory, "bookmarks_removed.pdf"));

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
        stream.Dispose();
    }

    static byte[] ReadAllBytes(Stream stream) {
        if(stream is MemoryStream memoryStream)
            return memoryStream.ToArray();

        using var copy = new MemoryStream();
        stream.Position = 0;
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    static Stream AddBookmarks(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        foreach(var page in processor.DocumentFacade.Pages) {
            var destination = page.CreateFitBBoxDestination();
            var bookmark = new PdfBookmark() {
                Destination = destination,
                Title = $"Page {destination.PageIndex + 1}"
            };
            processor.Document.Bookmarks.Add(bookmark);
        }

        return SaveToMemoryStream(processor);
    }

    static Stream DeleteBookmarks(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);
        processor.Document.Bookmarks.Clear();
        return SaveToMemoryStream(processor);
    }

    static MemoryStream SaveToMemoryStream(PdfDocumentProcessor processor) {
        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
