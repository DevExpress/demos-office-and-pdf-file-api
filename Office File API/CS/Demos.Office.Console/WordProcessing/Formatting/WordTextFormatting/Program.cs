using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Formatted_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FormatDocument(inputStream, OutputFormat);
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

    static Stream FormatDocument(Stream inputStream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        var range = document.Paragraphs[0].Range;

        var properties = document.BeginUpdateCharacters(range);
        properties.FontName = "Segoe UI";
        properties.FontSize = 16f;
        properties.ForeColor = Color.Black;
        properties.BackColor = ColorTranslator.FromHtml("#FFF7CC");
        properties.Underline = UnderlineType.Single;
        properties.Bold = true;
        properties.Italic = false;
        document.EndUpdateCharacters(properties);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
