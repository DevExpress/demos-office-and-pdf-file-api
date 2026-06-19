using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = RotateResizePages(File.OpenRead(inputPath), pageRange: [], rotate: 90, PdfPaperSize.A4);
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

    static Stream RotateResizePages(Stream documentStream, IEnumerable<int> pageRange, int rotate, PdfRectangle pageSize) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, processor.Document.Pages.Count);

        foreach(var index in pageRange) {
            if(index < 0 || index >= processor.Document.Pages.Count)
                continue;

            var page = processor.Document.Pages[index];
            page.Resize(pageSize, PdfContentHorizontalAlignment.Center, PdfContentVerticalAlignment.Center);
            page.Rotate = rotate;
        }

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
