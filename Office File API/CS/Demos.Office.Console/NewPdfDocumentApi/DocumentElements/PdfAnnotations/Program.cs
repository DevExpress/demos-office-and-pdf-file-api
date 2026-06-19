using System.Drawing;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const string SearchText = "DevExpress";
    const string FreeText = "Sample Text";
    const string AnnotationColor = "#CC5B00";
    const string StampIconName = "Approved";
    const string StickyNoteIconName = "Note";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);
            var color = ColorTranslator.FromHtml(AnnotationColor);

            SaveDocument(CreateLink(new MemoryStream(inputBytes), SearchText, color), Path.Combine(outputDirectory, "annotation_link.pdf"));
            SaveDocument(CreateFreeTextAnnotation(new MemoryStream(inputBytes), FreeText, color), Path.Combine(outputDirectory, "annotation_freetext.pdf"));
            SaveDocument(CreateRubberStamp(new MemoryStream(inputBytes), StampIconName), Path.Combine(outputDirectory, "annotation_stamp.pdf"));

            using var stickyNoteStream = CreateStickyNote(new MemoryStream(inputBytes), StickyNoteIconName, color);
            var stickyNoteBytes = ReadAllBytes(stickyNoteStream);
            SaveDocument(new MemoryStream(stickyNoteBytes), Path.Combine(outputDirectory, "annotation_stickynote.pdf"));
            SaveDocument(ClearAllAnnotations(new MemoryStream(stickyNoteBytes)), Path.Combine(outputDirectory, "annotation_cleared.pdf"));

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

    static Stream CreateLink(Stream documentStream, string searchText, Color color) {
        using var doc = new PdfDocument(documentStream);

        foreach(var result in doc.FindText(searchText)) {
            foreach(var match in result.Matches) {
                var bounds = match.GetRectangles().Select(x => x.BoundingBox)
                    .Aggregate(RectangleF.Union);
                var link = new LinkAnnotation(bounds) {
                    Action = new UriAction { Uri = "https://www.devexpress.com" },
                    Color = PdfColor.FromRgb(color.R, color.G, color.B),
                    BorderStyle = new AnnotationBorderStyle { Width = 2, Style = BorderStyle.Dashed }
                };
                doc.Pages[result.PageIndex].Annotations.Add(link);
            }
        }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream CreateFreeTextAnnotation(Stream documentStream, string text, Color color) {
        using var doc = new PdfDocument(documentStream);

        var page = doc.Pages[0];
        page.Annotations.Add(new FreeTextAnnotation(new RectangleF(450, 20, 130, 100)) {
            Content = text,
            Title = "DX Demo",
            Color = PdfColor.FromRgb(color.R, color.G, color.B),
            CreationDate = DateTimeOffset.UtcNow
        });

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream CreateRubberStamp(Stream documentStream, string iconName) {
        using var doc = new PdfDocument(documentStream);

        var page = doc.Pages[0];
        page.Annotations.Add(new RubberStampAnnotation(
            new RectangleF(350, 80, 230, 55),
            RubberStampAnnotationIconName.FromName(iconName)) {
            Title = "DX Demo",
            CreationDate = DateTimeOffset.UtcNow
        });

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream CreateStickyNote(Stream documentStream, string iconName, Color color) {
        using var doc = new PdfDocument(documentStream);

        var page = doc.Pages[0];
        page.Annotations.Add(new TextAnnotation(new RectangleF(450, 150, 20, 20)) {
            IconName = Enum.Parse<TextAnnotationIconName>(iconName),
            Color = PdfColor.FromRgb(color.R, color.G, color.B),
            Content = "This is a Sticky Note",
            Title = "DX Demo",
            CreationDate = DateTimeOffset.UtcNow
        });

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream ClearAllAnnotations(Stream documentStream) {
        using var doc = new PdfDocument(documentStream);

        foreach(var page in doc.Pages)
            page.Annotations.Clear();

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
