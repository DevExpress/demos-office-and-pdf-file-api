using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputDirectoryName = "Output";
    const string SearchText = "DevExpress";
    const string FreeText = "Sample Text";
    const string AnnotationColor = "#CC5B00";
    const string StampIconName = PdfRubberStampAnnotationIconName.Approved;
    const string StickyNoteIconName = PdfTextAnnotationIconName.Note;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);
            var color = System.Drawing.ColorTranslator.FromHtml(AnnotationColor);

            SaveDocument(CreateLink(new MemoryStream(inputBytes), SearchText, color), Path.Combine(outputDirectory, "annotation_link.pdf"));
            SaveDocument(CreateFreeTextAnnotation(new MemoryStream(inputBytes), FreeText, color), Path.Combine(outputDirectory, "annotation_freetext.pdf"));
            SaveDocument(CreateRubberStamp(new MemoryStream(inputBytes), StampIconName), Path.Combine(outputDirectory, "annotation_stamp.pdf"));

            using var stickyNoteStream = CreateStickyNote(new MemoryStream(inputBytes), StickyNoteIconName, color);
            var stickyNoteBytes = ReadAllBytes(stickyNoteStream);
            SaveDocument(new MemoryStream(stickyNoteBytes), Path.Combine(outputDirectory, "annotation_stickynote.pdf"));
            SaveDocument(FlattenAnnotations(new MemoryStream(stickyNoteBytes)), Path.Combine(outputDirectory, "annotation_flattened.pdf"));
            SaveDocument(ClearAnnotations(new MemoryStream(stickyNoteBytes)), Path.Combine(outputDirectory, "annotation_cleared.pdf"));

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

    static IReadOnlyList<PdfRectangle> FindText(PdfDocumentProcessor processor, string searchText) {
        var search = processor.FindText(searchText, new PdfTextSearchParameters() {
            CaseSensitive = false,
            WholeWords = false
        });

        return search.Status == PdfTextSearchStatus.Found
            ? search.Rectangles.Select(rectangle => rectangle.BoundingRectangle).ToList()
            : throw new InvalidOperationException($"Text '{searchText}' was not found in the source PDF.");
    }

    static Stream CreateLink(Stream documentStream, string searchText, System.Drawing.Color annotationColor) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var page = processor.DocumentFacade.Pages[0];
        var textRectangles = FindText(processor, searchText);

        foreach(var rectangle in textRectangles) {
            var linkAnnotation = page.AddLinkAnnotation(rectangle, "https://www.devexpress.com");
            linkAnnotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
            linkAnnotation.BorderWidth = 2;
            linkAnnotation.BorderStyle = PdfBorderStyle.Dash;
        }

        return SaveToMemoryStream(processor);
    }

    static Stream CreateFreeTextAnnotation(Stream documentStream, string text, System.Drawing.Color annotationColor) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var page = processor.DocumentFacade.Pages[0];
        var freeTextAnnotation = page.AddFreeTextAnnotation(new PdfRectangle(600, 10, 700, 110), text);
        freeTextAnnotation.Author = "DX Demo";
        freeTextAnnotation.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
        freeTextAnnotation.CreationDate = DateTime.UtcNow;

        return SaveToMemoryStream(processor);
    }

    static Stream CreateRubberStamp(Stream documentStream, string iconName) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var page = processor.DocumentFacade.Pages[0];
        var stamp = page.AddRubberStampAnnotation(new PdfPoint(600, 100), iconName);
        stamp.Author = "DX Demo";
        stamp.CreationDate = DateTime.UtcNow;

        return SaveToMemoryStream(processor);
    }

    static Stream CreateStickyNote(Stream documentStream, string iconName, System.Drawing.Color annotationColor) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var page = processor.DocumentFacade.Pages[0];
        var stickyNote = page.AddTextAnnotation(new PdfPoint(600, 100), iconName);
        stickyNote.Author = "DX Demo";
        stickyNote.Color = new PdfRGBColor(annotationColor.R / 255.0, annotationColor.G / 255.0, annotationColor.B / 255.0);
        stickyNote.CreationDate = DateTime.UtcNow;

        return SaveToMemoryStream(processor);
    }

    static Stream FlattenAnnotations(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);
        processor.DocumentFacade.FlattenAnnotations();
        return SaveToMemoryStream(processor);
    }

    static Stream ClearAnnotations(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var annotations = processor.DocumentFacade.Pages.SelectMany(page => page.Annotations).ToList();
        foreach(var annotation in annotations)
            annotation.Remove();

        return SaveToMemoryStream(processor);
    }

    static MemoryStream SaveToMemoryStream(PdfDocumentProcessor processor) {
        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
