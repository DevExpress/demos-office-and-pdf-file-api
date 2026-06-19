using DevExpress.XtraRichEdit;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "TrackChanges_template.docx");
    const string OutputFileName = "TrackChanges_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ProcessTrackChangesDocument(inputStream, OutputFormat);
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

    static Stream ProcessTrackChangesDocument(Stream inputStream, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        document.Revisions.AcceptAll();
        document.TrackChanges.Enabled = true;

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
