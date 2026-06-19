using DevExpress.Drawing;
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

            SaveDocument(AddPage(new MemoryStream(inputBytes), PdfPaperSize.A4), Path.Combine(outputDirectory, "add_page.pdf"));
            SaveDocument(DeletePages(new MemoryStream(inputBytes), [1]), Path.Combine(outputDirectory, "delete_page.pdf"));
            SaveDocument(DuplicatePages(new MemoryStream(inputBytes), [0]), Path.Combine(outputDirectory, "duplicate_page.pdf"));

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

    static Stream AddPage(Stream documentStream, PdfRectangle pageSize) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var newPage = processor.AddNewPage(pageSize);

        var text = "New Page";
        var font = new DXFont("Arial", 25);
        using var brush = new DXSolidBrush(System.Drawing.Color.Black);
        using var g = processor.CreateGraphicsPageSystem();

        var textSize = g.MeasureString(text, font);

        g.DrawString(text, font, brush, (float)(pageSize.Width - textSize.Width) / 2, (float)(pageSize.Height - textSize.Height) / 2);
        g.AddToPageForeground(newPage);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream DeletePages(Stream documentStream, IEnumerable<int> pageRange) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        foreach(var index in pageRange) {
            if(index < 1 || index > processor.Document.Pages.Count)
                continue;

            processor.DeletePage(index);
        }

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream DuplicatePages(Stream documentStream, IEnumerable<int> pageRange) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        foreach(var index in pageRange) {
            if(index < 0 || index >= processor.Document.Pages.Count)
                continue;

            processor.Document.Pages.Add(processor.Document.Pages[index]);
        }

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
