using DevExpress.Docs.Pdf;
using DevExpress.Drawing;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Alternative.pdf");
    const string OutputDirectoryName = "Output";
    const string ImageFormatName = "png";
    static readonly int[] PageRange = Array.Empty<int>();

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            using var documentStream = File.OpenRead(inputPath);
            var images = ExtractImages(documentStream, GetImageFormat(ImageFormatName), PageRange);
            for(int index = 0; index < images.Count; index++) {
                using var stream = images[index];
                SaveDocument(stream, Path.Combine(outputDirectory, $"image_{index + 1}.{ImageFormatName}"));
            }

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static void SaveDocument(Stream stream, string outputPath) {
        stream.Position = 0;
        using var fileStream = File.Create(outputPath);
        stream.CopyTo(fileStream);
    }

    static DXImageFormat GetImageFormat(string format) {
        return format.ToLowerInvariant() switch {
            "png" => DXImageFormat.Png,
            "jpeg" => DXImageFormat.Jpeg,
            "bmp" => DXImageFormat.Bmp,
            "tiff" => DXImageFormat.Tiff,
            "svg" => DXImageFormat.Svg,
            _ => DXImageFormat.Png
        };
    }

    static IReadOnlyList<Stream> ExtractImages(Stream documentStream, DXImageFormat imageFormat, IEnumerable<int> pageRange) {
        using var doc = new PdfDocument(documentStream);
        var result = new List<Stream>();

        if(!pageRange.Any())
            pageRange = Enumerable.Range(0, doc.Pages.Count);

        foreach(var index in pageRange) {
            if(index < 0 || index >= doc.Pages.Count)
                continue;

            foreach(var fragment in doc.Pages[index].Fragments) {
                if(fragment is ImageFragment imageFragment) {
                    var imageStream = new MemoryStream();
                    imageFragment.Image.Save(imageStream, imageFormat);
                    imageStream.Position = 0;
                    result.Add(imageStream);
                }
            }
        }

        return result;
    }
}
