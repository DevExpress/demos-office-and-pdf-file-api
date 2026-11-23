using DevExpress.Docs.Pdf;
using DevExpress.Drawing;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputFileName = "page_image";
    const string OutputDirectoryName = "Output";
    const string ImageFormat = "png";
    const int Resolution = 300;
    const RangeExportMode ExportMode = RangeExportMode.All;
    const string PageRange = "";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var documentStream = File.OpenRead(inputPath);

            var pageRange = RangeHelper.GetIndexesByRange(ExportMode, PageRange);
            var imageStreams = ExportToImages(documentStream, ImageFormat, Resolution, pageRange);

            if(imageStreams.Count == 0)
                throw new InvalidOperationException("Failed to export document pages to images.");

            if(imageStreams.Count == 1) {
                var outputPath = Path.Combine(outputDirectory, $"{OutputFileName}.{GetOutputExtension(ImageFormat)}");
                using var fileStream = File.Create(outputPath);
                using var imageStream = imageStreams[0];
                imageStream.CopyTo(fileStream);
                Console.WriteLine($"Created {outputPath}");
                return 0;
            }

            for(var index = 0; index < imageStreams.Count; index++) {
                var outputPath = Path.Combine(outputDirectory, $"page_{index + 1}.{GetOutputExtension(ImageFormat)}");
                using var fileStream = File.Create(outputPath);
                using var imageStream = imageStreams[index];
                imageStream.CopyTo(fileStream);
            }

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static IReadOnlyList<Stream> ExportToImages(Stream documentStream, string imageFormat, int resolution, IEnumerable<int> pageRange) {
        using var document = new PdfDocument(documentStream);

        var pageIndexes = pageRange
            .Where(index => index >= 0 && index < document.Pages.Count)
            .Distinct()
            .ToArray();

        if(!pageIndexes.Any())
            pageIndexes = Enumerable.Range(0, document.Pages.Count).ToArray();

        var exportOptions = new ImageExportOptions() {
            Resolution = resolution,
            ImageExportType = imageFormat == "svg" ? ImageExportType.Svg : ImageExportType.Bitmap
        };

        var images = document.ExportToImages(exportOptions, pageIndexes);
        var targetFormat = GetImageFormat(imageFormat);

        var result = new List<Stream>();
        foreach(var image in images) {
            var imageStream = new MemoryStream();
            image.Save(imageStream, targetFormat);
            imageStream.Position = 0;
            result.Add(imageStream);
        }

        return result;
    }

    static DXImageFormat GetImageFormat(string format) {
        return format.ToLowerInvariant() switch {
            "png" => DXImageFormat.Png,
            "jpeg" => DXImageFormat.Jpeg,
            "bmp" => DXImageFormat.Bmp,
            "tiff" => DXImageFormat.Tiff,
            "svg" => DXImageFormat.Svg,
            _ => throw new ArgumentException("Incorrect image format", nameof(format))
        };
    }

    static string GetOutputExtension(string imageFormat) {
        return imageFormat.ToLowerInvariant() switch {
            "bmp" => "bmp",
            "png" => "png",
            "jpeg" => "jpeg",
            "tiff" => "tiff",
            "svg" => "svg",
            _ => throw new ArgumentException("Incorrect image format", nameof(imageFormat))
        };
    }

    enum RangeExportMode {
        First,
        All,
        CustomRange
    }

    static class RangeHelper {
        public static List<int> GetIndexesByRange(RangeExportMode exportMode, string range) {
            return exportMode switch {
                RangeExportMode.First => new List<int> { 0 },
                RangeExportMode.All => new List<int>(),
                RangeExportMode.CustomRange => new List<int>(ParseRange(range)),
                _ => new List<int>()
            };
        }

        static IEnumerable<int> ParseRange(string range) {
            if(string.IsNullOrWhiteSpace(range))
                return Enumerable.Empty<int>();

            var result = new SortedSet<int>();
            var tokens = range.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);

            foreach(var raw in tokens) {
                var part = raw.Trim();
                if(part.Length == 0)
                    continue;

                var dashIndex = part.IndexOf('-');
                if(dashIndex > 0) {
                    var startStr = part.Substring(0, dashIndex).Trim();
                    var endStr = part.Substring(dashIndex + 1).Trim();

                    if(int.TryParse(startStr, out var start) && int.TryParse(endStr, out var end)) {
                        if(start <= 0 || end <= 0)
                            continue;

                        if(end < start)
                            (start, end) = (end, start);

                        for(var value = start; value <= end; value++)
                            result.Add(value - 1);
                    }

                    continue;
                }

                if(int.TryParse(part, out var single) && single > 0)
                    result.Add(single - 1);
            }

            return result;
        }
    }
}
