using DevExpress.Pdf;
using DevExpress.Pdf.Xmp;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    static readonly string MetadataFilePath = Path.Combine("Documents", "PDF", "Sample_Metadata.xmp");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var metadataPath = Path.Combine(AppContext.BaseDirectory, MetadataFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            SaveDocument(GetMetadata(File.OpenRead(inputPath)), Path.Combine(outputDirectory, "Metadata.xmp"));
            SaveDocument(RemoveMetadata(File.OpenRead(inputPath)), Path.Combine(outputDirectory, "metadata_removed.pdf"));
            SaveDocument(AddMetadata(File.OpenRead(inputPath), File.OpenRead(metadataPath)), Path.Combine(outputDirectory, "metadata_added.pdf"));

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

    static Stream GetMetadata(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var xmpDoc = XmpDocument.FromString(processor.Document.Metadata.Data);

        var outputStream = new MemoryStream();
        xmpDoc.Serialize(outputStream, new XmpSerializationOptions());

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream RemoveMetadata(Stream documentStream) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var xmpDoc = XmpDocument.FromString(processor.Document.Metadata.Data);
        xmpDoc.Clear();
        processor.Document.SetMetadata(xmpDoc);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream AddMetadata(Stream documentStream, Stream metadataStream) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var metadata = XmpDocument.FromStream(metadataStream);
        processor.Document.SetMetadata(metadata);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }
}
