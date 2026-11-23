using System.Text;
using System.Text.Json;
using DevExpress.Docs.Ocr;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "OCR", "Report_Scan.pdf");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "RecognizedRegion";
    const string OutputFormat = "pdf";
    const PageRegion Region = PageRegion.Header;

    static readonly Dictionary<PageRegion, System.Drawing.RectangleF> PredefinedRegions = new() {
        [PageRegion.Header] = new(0, 762, 595, 79),
        [PageRegion.Body] = new(0, 140, 595, 610),
        [PageRegion.Footer] = new(0, 0, 595, 128)
    };

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
            using var outputStream = RecognizeRegion(documentStream, OutputFormat);

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

    static MemoryStream RecognizeRegion(Stream documentStream, string outputFormat) {
        using var provider = new DocTrOcrProvider(new DocTrOcrProviderOptions { DetectPageOrientation = false });
        var engine = new OcrEngine(provider);

        using var document = new PdfDocument(documentStream);

        var options = new PdfDocumentRecognitionOptions {
            Regions = new[] { PredefinedRegions[Region] }
        };

        var result = engine.RecognizePdfDocument(document, options);

        var outputStream = new MemoryStream();
        if(outputFormat == "pdf") {
            using var pdf = result.ToPdf();
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
        var text = string.Join(Environment.NewLine, pages.SelectMany(page => page.TextFragments).Select(fragment => fragment.Text));
        return text.Length > 0 ? text : "Nothing found";
    }

    static string ToJson(IReadOnlyList<PdfPageRecognitionResult> pages) {
        var result = pages.Select(page => new OcrPage(page.PageIndex + 1, page.TextFragments.Select(fragment =>
            new OcrFragment(fragment.Text, Round(fragment.Confidence), ToBounds(fragment.Bounds),
                fragment.Words.Select(word => new OcrWord(word.Text, Round(word.Confidence), ToBounds(word.Bounds))).ToList())).ToList()));
        return JsonSerializer.Serialize(new { region = Region.ToString(), pages = result }, JsonOptions);
    }

    static OcrBounds ToBounds(System.Drawing.RectangleF bounds) {
        return new OcrBounds((float)Math.Round(bounds.X, 2), (float)Math.Round(bounds.Y, 2),
            (float)Math.Round(bounds.Width, 2), (float)Math.Round(bounds.Height, 2));
    }

    static float? Round(float? confidence) {
        return confidence.HasValue ? (float)Math.Round(confidence.Value, 3) : null;
    }

    enum PageRegion {
        Header,
        Body,
        Footer
    }

    record OcrBounds(float X, float Y, float Width, float Height);

    record OcrWord(string Text, float? Confidence, OcrBounds Bounds);

    record OcrFragment(string Text, float? Confidence, OcrBounds Bounds, IReadOnlyList<OcrWord> Words);

    record OcrPage(int Page, IReadOnlyList<OcrFragment> Fragments);
}
