using DevExpress.Data.Utils.Security;
using DevExpress.Drawing;
using DevExpress.Pdf;
using System.Drawing;

internal static class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "PDF", "Sample_Blank.pdf");
    static readonly string ImageFilePath = Path.Combine("Documents", "DevExpress.png");
    const string OutputDirectoryName = "Output";

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var imagePath = Path.Combine(AppContext.BaseDirectory, ImageFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);

            var inputBytes = File.ReadAllBytes(inputPath);

            SaveDocument(DrawPageText(new MemoryStream(inputBytes), "Sample Text"), Path.Combine(outputDirectory, "draw_text.pdf"));
            SaveDocument(DrawPageImage(new MemoryStream(inputBytes), File.OpenRead(imagePath)), Path.Combine(outputDirectory, "draw_image.pdf"));
            SaveDocument(DrawPageShapes(new MemoryStream(inputBytes), drawRectangle: true, drawEllipse: true, drawPolygon: true, drawLines: true, drawBeziers: true, drawPath: true), Path.Combine(outputDirectory, "draw_shapes.pdf"));

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

    static Stream DrawPageText(Stream documentStream, string textToDraw) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        using var g = processor.CreateGraphicsPageSystem();
        var page = processor.Document.Pages[0];

        using var brush = new DXSolidBrush(Color.Black);
        var font = new DXFont("Arial", 25);

        var stringSize = g.MeasureString(textToDraw, font);
        var x = (page.CropBox.Width - stringSize.Width) / 2;
        var y = (page.CropBox.Height - stringSize.Height) / 2;

        g.DrawString(textToDraw, font, brush, (float)x, (float)y);
        g.AddToPageForeground(page);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;

        return outputStream;
    }

    static Stream DrawPageImage(Stream documentStream, Stream imageStream) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        using var g = processor.CreateGraphicsPageSystem();
        var page = processor.Document.Pages[0];

        var image = DXImage.FromStream(imageStream);
        var rect = new RectangleF((float)page.CropBox.Left + 150, (float)page.CropBox.Height / 2 - 40, (float)page.CropBox.Width - 300, 40);

        g.DrawImage(image, rect);
        g.AddToPageForeground(page);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;

        return outputStream;
    }

    static Stream DrawPageShapes(Stream documentStream, bool drawRectangle, bool drawEllipse, bool drawPolygon, bool drawLines, bool drawBeziers, bool drawPath) {
        using var processor = new PdfDocumentProcessor();
        processor.LoadDocument(documentStream);

        using var g = processor.CreateGraphicsPageSystem();
        var page = processor.Document.Pages[0];

        using var brush = new DXSolidBrush(Color.Black);
        using var pen = new DXPen(brush);

        var pageCropBox = new RectangleF(0, 0, (float)page.CropBox.Width, (float)page.CropBox.Height);
        var cell = new RectangleF(0, 0, pageCropBox.Width / 2, pageCropBox.Height / 3);

        if(drawRectangle)
            DrawRectangle(g, pen, pageCropBox, cell);
        if(drawEllipse)
            DrawEllipse(g, pen, pageCropBox, cell);
        if(drawPolygon)
            DrawPolygon(g, pen, pageCropBox, cell);
        if(drawLines)
            DrawLines(g, pen, pageCropBox, cell);
        if(drawBeziers)
            DrawBeziers(g, pen, pageCropBox, cell);
        if(drawPath)
            DrawPath(g, pen, pageCropBox, cell);

        g.AddToPageForeground(page);

        var outputStream = new MemoryStream();
        processor.SaveDocument(outputStream);
        outputStream.Position = 0;

        return outputStream;
    }

    static void DrawRectangle(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var rectCell = new RectangleF(pageCropBox.Left + 10, pageCropBox.Height - cell.Height + 10, cell.Width - 20, cell.Height - 20);
        g.DrawRectangle(pen, rectCell);
    }

    static void DrawEllipse(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var ellipseCell = new RectangleF(pageCropBox.Left + 10, pageCropBox.Height - cell.Height * 2 + 10, cell.Width - 20, cell.Height - 20);
        g.DrawEllipse(pen, ellipseCell);
    }

    static void DrawPolygon(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var starCell = new RectangleF(pageCropBox.Left + 10, pageCropBox.Height - cell.Height * 3 + 10, cell.Width - 20, cell.Height - 20);
        g.DrawPolygon(pen, Enumerable.Range(0, 10).Select(x => {
            var angle = Math.PI / 2 + x * Math.PI / 5;
            var radius = starCell.Height / 2;
            var innerRadius = x % 2 == 0 ? radius : radius * ((3 - Math.Sqrt(5)) / 2);
            return new PointF((float)(starCell.Left + starCell.Width / 2 + innerRadius * Math.Cos(angle)),
                (float)(starCell.Top + starCell.Height / 2 + innerRadius * Math.Sin(angle)));
        }).ToArray());
    }

    static void DrawLines(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var linesCell = new RectangleF(pageCropBox.Left + cell.Width + 10, pageCropBox.Height - cell.Height + 10, cell.Width - 20, cell.Height - 20);
        g.DrawLines(pen, Enumerable.Range(0, 10).Select(x => new PointF(linesCell.Left + linesCell.Width / 10 * x,
            StrongRandom.Next((int)linesCell.Top, (int)linesCell.Bottom))).ToArray());
    }

    static void DrawBeziers(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var bezierCell = new RectangleF(pageCropBox.Left + cell.Width + 10, pageCropBox.Height - cell.Height * 2 + 10, cell.Width - 20, cell.Height - 20);
        g.DrawBeziers(pen, Enumerable.Range(0, 10).Select(x => new PointF(bezierCell.Left + bezierCell.Width / 10 * x,
            StrongRandom.Next((int)bezierCell.Top, (int)bezierCell.Bottom))).ToArray());
    }

    static void DrawPath(PdfGraphics g, DXPen pen, RectangleF pageCropBox, RectangleF cell) {
        var pathCell = new RectangleF(pageCropBox.Left + cell.Width + 10, pageCropBox.Height - cell.Height * 3 + 10, cell.Width - 20, cell.Height - 20);
        using var path = new DXGraphicsPath();
        path.AddEllipse(new(pathCell.Left, pathCell.Top, pathCell.Height, pathCell.Height));
        path.AddEllipse(new(pathCell.Left + pathCell.Height / 5, pathCell.Top + pathCell.Height * 2 / 3, pathCell.Height / 5, pathCell.Height / 5));
        path.AddEllipse(new(pathCell.Left + pathCell.Height * 3 / 5, pathCell.Top + pathCell.Height * 2 / 3, pathCell.Height / 5, pathCell.Height / 5));
        path.AddCurve([
            new(pathCell.Left + pathCell.Height / 6, pathCell.Top + pathCell.Height / 3),
            new(pathCell.Left + pathCell.Height / 2, pathCell.Top + pathCell.Height / 6),
            new(pathCell.Left + pathCell.Height * 5 / 6, pathCell.Top + pathCell.Height / 3)
        ], 0.7f);
        g.DrawPath(pen, path);
    }
}
