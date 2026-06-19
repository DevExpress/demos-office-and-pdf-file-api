using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputFileName = "ClearedPages.pdf";
    const bool ClearAnnotations = true;
    const bool ClearGraphics = true;
    const bool ClearImages = true;
    const bool ClearText = true;
    static readonly int[] PageRange = Array.Empty<int>();

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = ClearPages(documentStream, ClearAnnotations, ClearGraphics, ClearImages, ClearText, PageRange);
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

    static Stream ClearPages(Stream documentStream, bool clearAnnotations, bool clearGraphics, bool clearImages, bool clearText, IEnumerable<int> pageRange) {
        using var doc = new PdfDocument(documentStream);

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, doc.Pages.Count);

        foreach(var index in pageRange) {
            if(index < 0 || index >= doc.Pages.Count)
                continue;
            var clearRegion = new OrientedRectangle(doc.Pages[index].MediaBox);
            doc.ClearContent(index, [clearRegion], clearText, clearImages, clearGraphics, clearAnnotations);
        }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
