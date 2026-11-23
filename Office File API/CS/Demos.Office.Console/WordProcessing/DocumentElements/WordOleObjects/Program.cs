using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;

class Program {
    static readonly string OleObjectPath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    static readonly string IconPath = Path.Combine("Documents", "Excel.png");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "OleObject_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var oleObjectPath = Path.Combine(AppContext.BaseDirectory, OleObjectPath);
            var iconPath = Path.Combine(AppContext.BaseDirectory, IconPath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var oleObjectStream = File.OpenRead(oleObjectPath);
            using var iconStream = File.OpenRead(iconPath);
            using var outputStream = GenerateDocumentWithOleObject(oleObjectStream, iconStream, OutputFormat);
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

    static Stream GenerateDocumentWithOleObject(Stream oleObjectStream, Stream imageStream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;

        var oleShape = document.Shapes.InsertOleObjectAsIcon(document.Range.Start, oleObjectStream, OleObjectType.ExcelWorksheet, DocumentImageSource.FromStream(imageStream));
        oleShape.TextWrapping = TextWrappingType.InLineWithText;

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
