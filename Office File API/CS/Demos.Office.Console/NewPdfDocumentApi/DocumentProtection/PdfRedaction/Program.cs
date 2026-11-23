using System.Drawing;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "RedactedDocument.pdf";
    const string SearchText = "PDF Viewer";
    const string RedactionText = "REDACTED";
    const bool ApplyRedaction = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = AddRedaction(documentStream, SearchText, RedactionText, ApplyRedaction);
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

    static Stream AddRedaction(Stream documentStream, string searchText, string redactionText, bool applyRedaction) {
        using var doc = new PdfDocument(documentStream);

        var searchResults = doc.FindText(searchText, new TextSearchOptions());

        foreach(var result in searchResults) {
            List<RedactionAnnotation> annotations = new();
            foreach(var match in result.Matches) {
                var rects = match.GetRectangles().ToArray();
                var annotation = new RedactionAnnotation(RectangleF.Empty) {
                    Geometry = new RedactionGeometry(rects),
                    FillColor = PdfColor.Black,
                    Color = PdfColor.Red,
                    OverlayText = string.IsNullOrEmpty(redactionText) ? null : redactionText,
                    CreationDate = DateTime.UtcNow,
                    TextJustification = TextJustification.LeftJustified,
                    RepeatText = true,
                    TextAppearance = new TextAppearance() {
                        FontSize = 0,
                        Fill = SolidFill.White,
                    }
                };
                annotations.Add(annotation);
            }
            if(applyRedaction)
                doc.ApplyRedaction(result.PageIndex, annotations.ToArray());
            else
                foreach(var annotation in annotations)
                    doc.Pages[result.PageIndex].Annotations.Add(annotation);
        }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
