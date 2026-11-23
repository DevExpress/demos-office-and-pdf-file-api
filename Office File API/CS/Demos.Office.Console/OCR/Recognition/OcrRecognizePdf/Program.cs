using System.Text;
using System.Text.Json;
using DevExpress.Docs.Ocr;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "OCR", "Letter_Scan.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "RecognizedPdf";
    const string OutputFormat = "pdf";
    const PdfDocumentRecognitionMode RecognitionMode = PdfDocumentRecognitionMode.RenderedPage;

    static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = RecognizePages(documentStream, OutputFormat);

            var outputPath = Path.Combine(outputDirectory, $"{OutputFileName}.{OutputFormat}");
            File.WriteAllBytes(outputPath, outputStream.ToArray());

            Console.WriteLine($"Created {outputPath}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static MemoryStream RecognizePages(Stream documentStream, string outputFormat) {
        using var provider = new DocTrOcrProvider(new DocTrOcrProviderOptions { DetectPageOrientation = true });
        var engine = new OcrEngine(provider);

        using var document = new PdfDocument(documentStream);

        var options = new PdfDocumentRecognitionOptions { RecognitionMode = RecognitionMode };

        var result = engine.RecognizePdfDocument(document, options);

        var outputStream = new MemoryStream();
        if(outputFormat == "pdf") {
            using var pdf = result.ToPdf(new PdfGenerationOptions { RotatePages = true });
            pdf.Save(outputStream);
        }
        else {
            var pages = result.PageResults.ToList();
            var content = outputFormat == "json" ? ToJson(pages) : ToText(pages);
            outputStream.Write(new UTF8Encoding(false).GetBytes(content));
        }
        return outputStream;
    }

    static string ToText(IReadOnlyList<PdfPageRecognitionResult> pages) {
        var text = pages.Count == 1
            ? string.Join(Environment.NewLine, pages[0].TextFragments.Select(fragment => fragment.Text))
            : string.Join(Environment.NewLine + Environment.NewLine, pages.Select(page =>
                $"--- Page {page.PageIndex + 1} ---{Environment.NewLine}" +
                string.Join(Environment.NewLine, page.TextFragments.Select(fragment => fragment.Text))));
        return text.Length > 0 ? text : "Nothing found";
    }

    static string ToJson(IReadOnlyList<PdfPageRecognitionResult> pages) {
        var result = pages.Select(page => new OcrPage(page.PageIndex + 1, page.TextFragments.Select(fragment =>
            new OcrFragment(fragment.Text, Round(fragment.Confidence), ToBounds(fragment.Bounds),
                fragment.Words.Select(word => new OcrWord(word.Text, Round(word.Confidence), ToBounds(word.Bounds))).ToList())).ToList()));
        return JsonSerializer.Serialize(new { pages = result }, JsonOptions);
    }

    static OcrBounds ToBounds(System.Drawing.RectangleF bounds) {
        return new OcrBounds((float)Math.Round(bounds.X, 2), (float)Math.Round(bounds.Y, 2),
            (float)Math.Round(bounds.Width, 2), (float)Math.Round(bounds.Height, 2));
    }

    static float? Round(float? confidence) {
        return confidence.HasValue ? (float)Math.Round(confidence.Value, 3) : null;
    }

    record OcrBounds(float X, float Y, float Width, float Height);

    record OcrWord(string Text, float? Confidence, OcrBounds Bounds);

    record OcrFragment(string Text, float? Confidence, OcrBounds Bounds, IReadOnlyList<OcrWord> Words);

    record OcrPage(int Page, IReadOnlyList<OcrFragment> Fragments);
}
