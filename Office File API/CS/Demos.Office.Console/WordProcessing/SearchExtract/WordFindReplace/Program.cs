using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputFileName = "Modified_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FindReplace(inputStream, OutputFormat);
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

    static Stream FindReplace(Stream inputStream, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        document.AppendText("\n\nsome text\nsome text\nsome text");
        document.ReplaceAll("some text", "replacement text", SearchOptions.None);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
