using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.docx");
    const string OutputFileName = "Formatted_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
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
        wordProcessor.Unit = DevExpress.Office.DocumentUnit.Inch;
        wordProcessor.LoadDocument(inputStream);
        var document = wordProcessor.Document;
        var range = document.Paragraphs[0].Range;

        var paragraphProperties = document.BeginUpdateParagraphs(range);
        paragraphProperties.BackColor = ColorTranslator.FromHtml("#F2F8FF");
        paragraphProperties.Alignment = ParagraphAlignment.Left;
        paragraphProperties.LineSpacingType = ParagraphLineSpacing.Multiple;
        paragraphProperties.LineSpacingMultiplier = 1.15f;
        paragraphProperties.LeftIndent = 0.5f;
        paragraphProperties.WidowOrphanControl = true;
        paragraphProperties.KeepWithNext = false;
        paragraphProperties.KeepLinesTogether = false;
        SetParagraphBorder(paragraphProperties.Borders.HorizontalBorder, 2f, BorderLineStyle.Single, ColorTranslator.FromHtml("#4682B4"));
        SetParagraphBorder(paragraphProperties.Borders.BottomBorder, 2f, BorderLineStyle.Single, ColorTranslator.FromHtml("#4682B4"));
        SetParagraphBorder(paragraphProperties.Borders.TopBorder, 2f, BorderLineStyle.Single, ColorTranslator.FromHtml("#4682B4"));
        SetParagraphBorder(paragraphProperties.Borders.LeftBorder, 2f, BorderLineStyle.Single, ColorTranslator.FromHtml("#4682B4"));
        SetParagraphBorder(paragraphProperties.Borders.RightBorder, 2f, BorderLineStyle.Single, ColorTranslator.FromHtml("#4682B4"));
        document.EndUpdateParagraphs(paragraphProperties);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void SetParagraphBorder(ParagraphBorder border, float lineWidth, BorderLineStyle lineStyle, Color borderColor) {
        border.LineWidth = lineWidth;
        border.LineStyle = lineStyle;
        border.LineColor = borderColor;
    }
}
