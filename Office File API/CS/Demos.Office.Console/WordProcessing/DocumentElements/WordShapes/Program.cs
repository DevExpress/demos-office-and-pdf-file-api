using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputImagePath = Path.Combine("Documents", "Photo1.jpeg");
    const string OutputFileName = "Shape_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const float ImageScale = 1f;
    const float OutlineWidth = 1f;
    static readonly Color OutlineColor = ColorTranslator.FromHtml("#9E67FF");

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputImagePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var imageStream = File.OpenRead(inputPath);
            using var outputStream = GenerateDocumentWithImage(imageStream, OutputFormat);
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

    static Stream GenerateDocumentWithImage(Stream imageStream, DocumentFormat outputFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;
        document.Unit = DevExpress.Office.DocumentUnit.Inch;

        var shape = document.Shapes.InsertPicture(document.Range.End, DocumentImageSource.FromStream(imageStream));
        shape.LockAspectRatio = false;
        shape.TextWrapping = TextWrappingType.InLineWithText;
        shape.ScaleX = ImageScale;
        shape.ScaleY = ImageScale;
        shape.Line.Color = OutlineColor;
        shape.Line.Thickness = OutlineWidth;

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
