using DevExpress.Docs.Presentation;
using DevExpress.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "Sample.pptx");
    const string OutputDirectoryName = "Output";
    static readonly DXImageFormat ImageFormat = DXImageFormat.Png;
    const int Resolution = 300;
    const RangeExportMode ExportMode = RangeExportMode.All;
    const string SlideRange = "";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var inputStream = File.OpenRead(inputPath);
            var slideIndexes = RangeHelper.GetIndexesByRange(ExportMode, SlideRange);
            var imageStreams = ConvertToImages(inputStream, ImageFormat, Resolution, slideIndexes);
            if(imageStreams.Count == 0)
                throw new InvalidOperationException("Failed to convert presentation to images.");

            for(int i = 0; i < imageStreams.Count; i++) {
                using var imageStream = imageStreams[i];
                var outputPath = Path.Combine(outputDirectory, $"slide_{i + 1}.{GetImageExtension(ImageFormat)}");
                using var fileStream = File.Create(outputPath);
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

    static IReadOnlyList<Stream> ConvertToImages(Stream inputStream, DXImageFormat imageFormat, int resolution, List<int> slideIndexes) {
        var outputStreams = new List<Stream>();
        using var presentation = new Presentation(inputStream);
        var images = presentation.ExportToImages(new ImageExportOptions { Resolution = resolution }, slideIndexes.ToArray());
        foreach(var image in images) {
            var outputStream = new MemoryStream();
            image.Save(outputStream, imageFormat);
            outputStream.Position = 0;
            outputStreams.Add(outputStream);
        }
        return outputStreams.AsReadOnly();
    }

    static string GetImageExtension(DXImageFormat imageFormat) {
        if(imageFormat.Equals(DXImageFormat.Bmp))
            return "bmp";
        if(imageFormat.Equals(DXImageFormat.Jpeg))
            return "jpeg";
        if(imageFormat.Equals(DXImageFormat.Svg))
            return "svg";
        return "png";
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

                int dashIndex = part.IndexOf('-');
                if(dashIndex > 0) {
                    var startText = part.Substring(0, dashIndex).Trim();
                    var endText = part[(dashIndex + 1)..].Trim();
                    if(int.TryParse(startText, out int start) && int.TryParse(endText, out int end)) {
                        if(start <= 0 || end <= 0)
                            continue;
                        if(end < start)
                            (start, end) = (end, start);
                        for(int value = start; value <= end; value++)
                            result.Add(value - 1);
                    }
                    continue;
                }

                if(int.TryParse(part, out int single) && single > 0)
                    result.Add(single - 1);
            }

            return result;
        }
    }
}
