using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const float NewPageWidth = 595.28f;
    const float NewPageHeight = 841.89f;
    static readonly int[] PageRange = { 0 };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);
            SaveDocument(AddPage(new MemoryStream(inputBytes), NewPageWidth, NewPageHeight), Path.Combine(outputDirectory, "page_added.pdf"));
            SaveDocument(DeletePages(new MemoryStream(inputBytes), PageRange), Path.Combine(outputDirectory, "page_deleted.pdf"));
            SaveDocument(DuplicatePages(new MemoryStream(inputBytes), PageRange), Path.Combine(outputDirectory, "page_duplicated.pdf"));

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

    static Stream AddPage(Stream documentStream, float width, float height) {
        using var doc = new PdfDocument(documentStream);
        doc.Pages.Add(width, height);
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream DeletePages(Stream documentStream, IEnumerable<int> pageRange) {
        using var doc = new PdfDocument(documentStream);
        foreach(var index in pageRange) {
            if(index < 0 || index >= doc.Pages.Count)
                continue;
            doc.Pages.RemoveAt(index);
        }
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream DuplicatePages(Stream documentStream, IEnumerable<int> pageRange) {
        using var doc = new PdfDocument(documentStream);
        foreach(var index in pageRange) {
            if(index < 0 || index >= doc.Pages.Count)
                continue;
            doc.Pages.Add(doc.Pages[index].Clone());
        }
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
