using System.Drawing;
using DevExpress.Docs.Pdf;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Main.pdf");
    const string OutputFileName = "PdfWithWatermark.pdf";
    const string WatermarkText = "WATERMARK";
    const bool Rotate = false;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var documentStream = File.OpenRead(inputPath);
            using var outputStream = ApplyWatermark(documentStream, WatermarkText, Rotate);
            SaveDocument(outputStream, outputPath);

            Console.WriteLine($"Created {outputPath}");
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

    static Stream ApplyWatermark(Stream documentStream, string text, bool rotate) {
        using var doc = new PdfDocument(documentStream);

        FormTemplate form = new() { Bounds = new RectangleF(0, 0, 842, 595) };
        form.AddFragment(new TextFragment() {
            Text = text,
            Location = rotate ? new PointF(300, 220) : new PointF(280, 300),
            ScaleX = 4,
            ScaleY = 4,
            ForegroundFill = new SolidFill(PdfColor.LightSkyBlue, 0.5),
            RotationAngle = rotate ? 45 : 0
        });

        foreach(var page in doc.Pages) {
            WatermarkAnnotation watermark = new(page.MediaBox) {
                Appearance = new AnnotationAppearances {
                    Normal = new AnnotationAppearance {
                        DefaultForm = form
                    }
                }
            };
            page.Annotations.Add(watermark);
        }

        var outputStream = new MemoryStream();
        doc.Save(outputStream);
        outputStream.Position = 0;
        return outputStream;
    }
}
