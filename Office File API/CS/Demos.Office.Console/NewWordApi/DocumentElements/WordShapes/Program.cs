using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using DevExpress.Drawing;
using System.Drawing;

class Program {
    static readonly string InputImagePath = Path.Combine("Documents", "Photo1.jpeg");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Shape_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const float ImageScale = 1f;
    const bool LockAspectRatio = false;
    const float OutlineWidth = 1f;
    static readonly Color OutlineColor = ColorTranslator.FromHtml("#9E67FF");

    const float DocumentUnitsPerInch = 300f;
    const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputImagePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var imageStream = File.OpenRead(inputPath);
            using var outputStream = GenerateDocumentWithImage(imageStream, ImageScale, LockAspectRatio, OutlineColor, OutlineWidth, OutputFormat);
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

    static Stream GenerateDocumentWithImage(Stream imageStream, float scaleFactor, bool lockAspectRatio,
                                            Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var document = new WordDocument();

        var picture = new PictureShape(DXImage.FromStream(imageStream));
        document.Body.AddParagraph().InlineElements.Add(picture);

        picture.LockAspectRatio = lockAspectRatio;
        picture.Width = picture.OriginalSize.Width * scaleFactor;
        picture.Height = picture.OriginalSize.Height * scaleFactor;
        picture.Outline = new LineStyle {
            Fill = new SolidFill(outlineColor),
            Width = outlineWidth * DocumentUnitsPerPoint
        };

        // Create output stream and save document in target format
        return CreateOutputStream(document, outputFormat);
    }

    static Stream CreateOutputStream(WordDocument document, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        // Save document in target format
        document.Save(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }
}