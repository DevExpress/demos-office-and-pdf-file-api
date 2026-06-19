using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var outputStream = FindHighlightText(
                new MemoryStream(File.ReadAllBytes(inputPath)),
                searchText: "DevExpress",
                annotationType: PdfTextMarkupAnnotationType.Highlight,
                annotationColor: System.Drawing.ColorTranslator.FromHtml("#CC5B00"));
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

    static Stream FindHighlightText(MemoryStream documentStream, string searchText, PdfTextMarkupAnnotationType annotationType, System.Drawing.Color annotationColor) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        PdfTextSearchResults search;

        do {
            search = processor.FindText(
                searchText,
                new PdfTextSearchParameters() {
                    CaseSensitive = false,
                    WholeWords = false
                });
            if(search.Status == PdfTextSearchStatus.Found) {
                var annotation = processor.DocumentFacade.Pages[search.Page.GetPageIndex()].AddTextMarkupAnnotation(search.Rectangles, annotationType);

                annotation.Author = "DX Demo";
                annotation.CreationDate = DateTime.UtcNow;
                annotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
            }
        } while(search.Status == PdfTextSearchStatus.Found);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
