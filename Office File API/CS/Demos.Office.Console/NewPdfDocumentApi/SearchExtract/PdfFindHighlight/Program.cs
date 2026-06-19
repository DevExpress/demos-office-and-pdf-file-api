using System.Drawing;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputFileName = "HighlightedDocument.pdf";
    const string SearchText = "DevExpress";
    const string AnnotationColor = "#FF0000";
    const TextMarkupAnnotationType AnnotationType = TextMarkupAnnotationType.Highlight;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            var color = ColorTranslator.FromHtml(AnnotationColor);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = FindHighlightText(documentStream, SearchText, color, AnnotationType);
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

    static Stream FindHighlightText(Stream documentStream, string searchText, Color color, TextMarkupAnnotationType annotationType) {
        using var doc = new PdfDocument(documentStream);

        foreach(var result in doc.FindText(searchText)) {
            foreach(var match in result.Matches) {
                var annotation = new TextMarkupAnnotation(match.GetRectangles().ToArray()) {
                    Color = PdfColor.FromRgb(color.R, color.G, color.B),
                    CreationDate = DateTime.UtcNow,
                    Title = "DX Demo",
                    Type = annotationType
                };
                doc.Pages[result.PageIndex].Annotations.Add(annotation);
            }
        }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
