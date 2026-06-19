using System.Text;
using DevExpress.Docs.Pdf;

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

            var inputBytes = File.ReadAllBytes(inputPath);
            File.WriteAllText(Path.Combine(outputDirectory, "Metadata.txt"), ReadMetadata(new MemoryStream(inputBytes)));
            SaveDocument(RemoveXmpMetadata(new MemoryStream(inputBytes)), Path.Combine(outputDirectory, "MetadataRemoved.pdf"));
            using var metadataStream = File.OpenRead(metadataPath);
            SaveDocument(AddXmpMetadata(new MemoryStream(inputBytes), metadataStream), Path.Combine(outputDirectory, "MetadataAdded.pdf"));

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

    static string ReadMetadata(Stream documentStream) {
        using var doc = new PdfDocument(documentStream);
        return doc.Metadata.Xmp?.ToXmlString() ?? string.Empty;
    }

    static Stream RemoveXmpMetadata(Stream documentStream) {
        using var doc = new PdfDocument(documentStream);
        doc.Metadata.Xmp = null;
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }

    static Stream AddXmpMetadata(Stream documentStream, Stream metadataStream) {
        using var doc = new PdfDocument(documentStream);
        doc.Metadata.Xmp = XmpMetadata.FromStream(metadataStream);
        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
