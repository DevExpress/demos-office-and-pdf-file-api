using DevExpress.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Form.pdf");
    static readonly string FormDataFilePath = Path.Combine("Documents", "PDF", "Sample_FormData.xml");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var formDataPath = Path.Combine(AppContext.BaseDirectory, FormDataFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var importedXml = ImportFormData(File.OpenRead(inputPath), File.OpenRead(formDataPath), PdfFormDataFormat.Xml);
            SaveDocument(importedXml, Path.Combine(outputDirectory, "Document.pdf"));

            using var exportedXml = ExportFormData(File.OpenRead(inputPath), PdfFormDataFormat.Xml);
            SaveDocument(exportedXml, Path.Combine(outputDirectory, "Data.xml"));

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

    static Stream ImportFormData(Stream documentStream, Stream formDataStream, PdfFormDataFormat formDataFormat) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);
        processor.ApplyFormData(new PdfFormData(formDataStream, formDataFormat));

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);

        outputStream.Position = 0;
        return outputStream;
    }

    static Stream ExportFormData(Stream documentStream, PdfFormDataFormat formDataFormat) {
        using var processor = new PdfDocumentProcessor();

        processor.LoadDocument(documentStream);

        var formData = processor.GetFormData();
        var outputStream = new MemoryStream();
        formData.Save(outputStream, formDataFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
