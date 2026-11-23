using DevExpress.Docs.Pdf;
using DevExpress.Docs.Word;
using DevExpress.Docs.Word.Export;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "ConvertedDocument.pdf";
    const PdfACompatibility PdfA = PdfACompatibility.None;
    const PdfUACompatibility PdfUA = PdfUACompatibility.None;
    const string Title = "Custom Title";
    const string Author = "Custom Author";
    const string Subject = "";
    const string Keywords = "";
    const bool ExportControlsAsAcroFormFields = false;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ConvertToPdf(
                inputStream,
                PdfA,
                PdfUA,
                Author,
                Title,
                Subject,
                Keywords,
                ExportControlsAsAcroFormFields);
            using var fileStream = File.Create(outputPath);
            outputStream.CopyTo(fileStream);

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static Stream ConvertToPdf(Stream inputStream, PdfACompatibility pdfAConformance, PdfUACompatibility pdfUACompatibility,
                               string author, string title, string subject, string keywords, bool exportControlsAsAcroForm) {
        // Load the input document
        using var document = new WordDocument(inputStream);

        // Create output stream
        var outputStream = new MemoryStream();

        // Export to PDF with specified settings
        var exportOptions = new PdfExportOptions() {
            PdfACompatibility = pdfAConformance,
            PdfUACompatibility = pdfUACompatibility,
            ExportInteractiveControls = exportControlsAsAcroForm
        };
        exportOptions.DocumentOptions.Author = author;
        exportOptions.DocumentOptions.Title = title;
        exportOptions.DocumentOptions.Subject = subject;
        exportOptions.DocumentOptions.Keywords = keywords;

        document.ExportToPdf(outputStream, exportOptions);

        outputStream.Position = 0;
        return outputStream;
    }
}