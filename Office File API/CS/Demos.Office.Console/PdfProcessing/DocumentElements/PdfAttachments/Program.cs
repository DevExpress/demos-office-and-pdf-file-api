using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";
    const string AttachmentFileName = "DevExpressLogo.png";
    const string AttachmentDescription = "Embedded sample image";
    const string AttachmentMimeType = "image/png";
    const string AttachmentPngBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y1koX8AAAAASUVORK5CYII=";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);
            var attachmentBytes = Convert.FromBase64String(AttachmentPngBase64);

            SaveDocument(AttachFile(new MemoryStream(inputBytes), AttachmentFileName, AttachmentDescription, AttachmentMimeType, attachmentBytes, false), Path.Combine(outputDirectory, "attachment_embedded.pdf"));
            SaveDocument(AttachFile(new MemoryStream(inputBytes), AttachmentFileName, AttachmentDescription, AttachmentMimeType, attachmentBytes, true), Path.Combine(outputDirectory, "attachment_annotation.pdf"));

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

    static Stream AttachFile(Stream documentStream, string fileName, string fileDescription, string fileMimeType, byte[] fileData, bool createAnnotation) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        var attachment = new PdfFileAttachment() {
            CreationDate = DateTime.UtcNow,
            FileName = fileName,
            Data = fileData,
            MimeType = fileMimeType,
            Description = fileDescription,
            ModificationDate = DateTime.UtcNow,
            Relationship = PdfAssociatedFileRelationship.Supplement
        };

        if(createAnnotation) {
            var page = processor.Document.Pages[0];
            var pageFacade = processor.DocumentFacade.Pages[0];
            pageFacade.AddFileAttachmentAnnotation(new PdfPoint(page.MediaBox.Right - 50, page.MediaBox.Bottom + 50), attachment, PdfFileAttachmentAnnotationIconName.PaperClip);
        }
        else {
            processor.AttachFile(attachment);
        }

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
