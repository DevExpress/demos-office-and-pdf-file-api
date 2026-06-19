using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Watermark_template.docx");
    const string OutputFileName = "Watermarked_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ApplyTextWatermark(inputStream, OutputFormat);
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

    static Stream ApplyTextWatermark(Stream inputStream, DocumentFormat targetFormat) {
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(inputStream);
        var options = new TextWatermarkOptions {
            FontFamily = "Calibri",
            FontSize = 72,
            Color = Color.Gray,
            Layout = WatermarkLayout.Diagonal,
            Semitransparent = true
        };
        wordProcessor.Document.WatermarkManager.SetText("CONFIDENTIAL", options);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
