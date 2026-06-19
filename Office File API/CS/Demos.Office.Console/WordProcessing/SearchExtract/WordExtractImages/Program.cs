using DevExpress.Drawing;
using DevExpress.Office.Services;
using DevExpress.Spreadsheet.Charts;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraSpreadsheet.Services;
using System.Drawing;

class Program {
    static readonly string InputImagePath = Path.Combine("Documents", "Photo1.jpeg");
    static readonly string ChartTemplatePath = Path.Combine("Documents", "ChartTemplate.crtx");
    const string OutputDirectoryName = "Document_Images";
    static readonly DXImageFormat ImageFormat = DXImageFormat.Png;
    const bool ExportCharts = true;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            if(Directory.Exists(outputDirectory))
                Directory.Delete(outputDirectory, true);
            Directory.CreateDirectory(outputDirectory);

            using var imageStream = File.OpenRead(Path.Combine(AppContext.BaseDirectory, InputImagePath));
            using var chartTemplateStream = File.OpenRead(Path.Combine(AppContext.BaseDirectory, ChartTemplatePath));
            using var documentStream = CreateSampleDocument(imageStream, chartTemplateStream);
            var imageStreams = ExtractImages(documentStream, ExportCharts, ImageFormat);
            if(imageStreams.Count == 0)
                throw new InvalidOperationException("Failed to extract document images.");

            for(int index = 0; index < imageStreams.Count; index++) {
                var outputPath = Path.Combine(outputDirectory, $"image_{index + 1}.{GetOutputExtension(ImageFormat)}");
                using var fileStream = File.Create(outputPath);
                using var outputStream = imageStreams[index];
                outputStream.CopyTo(fileStream);
            }

            Console.WriteLine($"Created {outputDirectory}");
            return 0;
        }
        catch(Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    static MemoryStream CreateSampleDocument(Stream imageStream, Stream chartTemplateStream) {
        OfficeCharts.Instance.ActivateCrossPlatformCharts();
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;
        document.Unit = DevExpress.Office.DocumentUnit.Inch;
        document.AppendText("Sample document used to extract images.\n\n");
        document.Shapes.InsertPicture(document.Range.End, DocumentImageSource.FromStream(imageStream)).TextWrapping = TextWrappingType.InLineWithText;
        var chartShape = document.Shapes.InsertChart(document.Range.End, DevExpress.XtraRichEdit.API.Native.ChartType.ColumnClustered);
        chartShape.Size = new SizeF(4, 3);
        chartShape.TextWrapping = TextWrappingType.InLineWithText;
        var chart = (ChartObject)chartShape.ChartFormat.Chart;
        var worksheet = (DevExpress.Spreadsheet.Worksheet)chartShape.ChartFormat.Worksheet;
        worksheet["A1"].Value = "Month";
        worksheet["B1"].Value = "Sales";
        worksheet["A2"].Value = "January";
        worksheet["B2"].Value = 12;
        worksheet["A3"].Value = "February";
        worksheet["B3"].Value = 17;
        worksheet["A4"].Value = "March";
        worksheet["B4"].Value = 9;
        chart.SelectData(worksheet.GetDataRange());
        chart.LoadTemplate(chartTemplateStream);

        var stream = new MemoryStream();
        wordProcessor.SaveDocument(stream, DocumentFormat.Docx);
        stream.Position = 0;
        return stream;
    }

    static IReadOnlyList<Stream> ExtractImages(Stream inputStream, bool exportCharts, DXImageFormat imageFormat) {
        OfficeCharts.Instance.ActivateCrossPlatformCharts();
        using var wordProcessor = new RichEditDocumentServer();
        var imageStreams = new List<Stream>();
        wordProcessor.LoadDocument(inputStream);
        foreach(var shape in wordProcessor.Document.Shapes) {
            if(shape.Type == ShapeType.Picture) {
                var stream = new MemoryStream();
                shape.PictureFormat.Picture.DXImage.Save(stream, imageFormat);
                stream.Position = 0;
                imageStreams.Add(stream);
            }
            else if(exportCharts && shape.Type == ShapeType.Chart) {
                var stream = new MemoryStream();
                ((ChartObject)shape.ChartFormat.Chart).GetImage().DXImage.Save(stream, imageFormat);
                stream.Position = 0;
                imageStreams.Add(stream);
            }
        }
        return imageStreams.AsReadOnly();
    }

    static string GetOutputExtension(DXImageFormat imageFormat) {
        if(imageFormat == DXImageFormat.Png)
            return "png";
        if(imageFormat == DXImageFormat.Jpeg)
            return "jpeg";
        if(imageFormat == DXImageFormat.Bmp)
            return "bmp";
        throw new ArgumentException("Unsupported image format.", nameof(imageFormat));
    }
}
