using System.Text;
using System.Text.Json;
using DevExpress.Docs.Ocr;
using DevExpress.Drawing;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "OCR", "Letter_Scan.jpg");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "RecognizedImage";
    const string OutputFormat = "pdf";

    static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase
    };

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var imageStream = File.OpenRead(inputPath);
            using var outputStream = RecognizeImage(imageStream, OutputFormat);

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

    static MemoryStream RecognizeImage(Stream imageStream, string outputFormat) {
        using var provider = new DocTrOcrProvider(new DocTrOcrProviderOptions { DetectPageOrientation = true });
        var engine = new OcrEngine(provider);

        using var image = DXImage.FromStream(imageStream);

        var results = engine.RecognizeImages(new[] { image });

        var outputStream = new MemoryStream();
        if(outputFormat == "pdf") {
            using var document = results.ToPdf(new PdfGenerationOptions { RotatePages = true });
            document.Save(outputStream);
        }
        else {
            var fragments = results.ImageResults.Single().TextFragments;
            var content = outputFormat == "json" ? ToJson(fragments) : ToText(fragments);
            outputStream.Write(new UTF8Encoding(false).GetBytes(content));
        }
        return outputStream;
    }

    static string ToText(IEnumerable<ImageTextFragment> fragments) {
        var text = string.Join(Environment.NewLine, fragments.Select(fragment => fragment.Text));
        return text.Length > 0 ? text : "Nothing found";
    }

    static string ToJson(IEnumerable<ImageTextFragment> fragments) {
        var result = fragments.Select(fragment => new OcrFragment(fragment.Text, Round(fragment.Confidence), ToBounds(fragment.Bounds),
            fragment.Words.Select(word => new OcrWord(word.Text, Round(word.Confidence), ToBounds(word.Bounds))).ToList()));
        return JsonSerializer.Serialize(new { fragments = result }, JsonOptions);
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
}
