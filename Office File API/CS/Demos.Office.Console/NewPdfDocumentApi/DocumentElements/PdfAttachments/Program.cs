using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    static readonly string AttachmentFilePath = Path.Combine("Documents", "DevExpress.png");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "PdfWithAttachment.pdf";
    const string AttachmentDescription = "DevExpress Logo Image";
    const string AttachmentMimeType = "image/png";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var attachmentPath = Path.Combine(AppContext.BaseDirectory, AttachmentFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            var attachmentBytes = File.ReadAllBytes(attachmentPath);
            using var outputStream = AttachFile(documentStream, Path.GetFileName(attachmentPath), AttachmentDescription, AttachmentMimeType, attachmentBytes);
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

    static Stream AttachFile(Stream documentStream, string fileName, string fileDescription, string fileMimeType, byte[] fileData) {
        using var doc = new PdfDocument(documentStream);
        var attachment = new Attachment() {
            CreationDate = DateTime.UtcNow,
            FileName = fileName,
            Data = fileData,
            MimeType = fileMimeType,
            Description = fileDescription,
            ModificationDate = DateTime.UtcNow,
            Relationship = AssociatedFileRelationship.Supplement
        };
        doc.Attachments.Add(attachment);
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
