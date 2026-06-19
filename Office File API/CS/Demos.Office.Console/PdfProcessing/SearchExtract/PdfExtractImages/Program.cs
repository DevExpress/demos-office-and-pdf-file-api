using DevExpress.Drawing;
using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var images = ExtractImages(File.OpenRead(inputPath), DXImageFormat.Png, 300, []);
            for(var i = 0; i < images.Count; i++) {
                SaveDocument(images[i], Path.Combine(outputDirectory, $"image_{i + 1}.png"));
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

    static IReadOnlyList<Stream> ExtractImages(Stream documentStream, DXImageFormat imageFormat, int resolution, IEnumerable<int> pageRange) {
        using var processor = new PdfDocumentProcessor();
        var result = new List<Stream>();

        processor.LoadDocument(documentStream);

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, processor.Document.Pages.Count);

        foreach(var index in pageRange) {
            if(index < 0 || index >= processor.Document.Pages.Count)
                continue;

            var page = processor.Document.Pages[index];

            var images = processor.GetDXImages(
                new PdfDocumentArea(page.GetPageIndex(), page.MediaBox),
                resolution);

            foreach(var image in images) {
                var imageStream = new MemoryStream();
                image.Save(imageStream, imageFormat);
                imageStream.Position = 0;
                result.Add(imageStream);
            }
        }

        return result;
    }
}
