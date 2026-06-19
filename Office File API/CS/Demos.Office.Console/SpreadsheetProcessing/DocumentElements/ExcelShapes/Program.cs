using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    const string OutputFileName = "Shape_Document.xlsx";
    const ShapeActionType ShapeType = ShapeActionType.Picture;
    const bool Embedded = true;
    const ShapeGeometryPreset ShapeGeometry = ShapeGeometryPreset.Rectangle;
    const string ShapeText = "Sample Text";
    const float OutlineWidth = 1.0f;
    const int NumberOfShapes = 3;
    static readonly Color OutlineColor = ColorTranslator.FromHtml("#9E67FF");
    static readonly Color FillColor = ColorTranslator.FromHtml("#FFFFFF");
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = ShapeType switch {
                ShapeActionType.Picture => GenerateWorkbookWithImage(CreateSampleImageStream(), Embedded, false, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Textbox => GenerateWorkbookWithTextbox(ShapeText, 1.5f, 1.5f, false, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Figure => GenerateWorkbookWithFigure(ShapeGeometry, 1.5f, 1.5f, false, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Group => GenerateWorkbookWithShapeGroup(ShapeGeometry, NumberOfShapes, 1.5f, 1.5f, false, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                _ => GenerateWorkbookWithImage(CreateSampleImageStream(), Embedded, false, OutlineColor, OutlineWidth, OutputFormat)
            };
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

    static Stream CreateSampleImageStream() {
        var stream = new MemoryStream();
        using var bitmap = new DXBitmap(160, 120);
        using(DXGraphics graphics = DXGraphics.FromImage(bitmap)) {
            graphics.Clear(Color.White);
            using var brush = new DXSolidBrush(ColorTranslator.FromHtml("#9E67FF"));
            graphics.FillRectangle(brush, 10, 10, 140, 100);
            graphics.DrawEllipse(DXPens.White, new RectangleF(45, 25, 70, 70));
        }
        bitmap.Save(stream, DXImageFormat.Png);
        stream.Position = 0;
        return stream;
    }

    static Stream GenerateWorkbookWithImage(Stream imageStream, bool embedded, bool lockAspectRatio, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Images";

        if(embedded) {
            var cell = worksheet.Cells["B2"];
            cell.Value = imageStream;
            if(cell.Value.IsCellImage) {
                cell.ImageInfo.Decorative = true;
                cell.ImageInfo.AlternativeText = "Sample Image";
            }
        }
        else {
            var imageSource = SpreadsheetImageSource.FromStream(imageStream);
            var imageShape = worksheet.Pictures.AddPicture(imageSource, worksheet.Range["B2:D11"]);
            imageShape.LockAspectRatio = lockAspectRatio;
            imageShape.Outline.SetSolidFill(outlineColor);
            imageShape.Outline.Width = outlineWidth;
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithTextbox(string text, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "TextBox";
        var textBox = worksheet.Shapes.AddTextBox(worksheet.Range["B2:D3"], text);
        FormatShape(textBox, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithFigure(ShapeGeometryPreset geometry, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Figure";
        var figure = worksheet.Shapes.AddShape(geometry, worksheet.Range["B2:D3"]);
        FormatShape(figure, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithShapeGroup(ShapeGeometryPreset geometry, int numberOfShapes, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Shape Group";
        Shape[] shapes = new Shape[numberOfShapes];
        for(int i = 0; i < numberOfShapes; i++) {
            shapes[i] = worksheet.Shapes.AddShape(geometry, i * 0.5f, i * 0.5f, width, height);
            FormatShape(shapes[i], width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
        }
        worksheet.Shapes.GroupShapes(shapes);
        return CreateOutputStream(workbook, outputFormat);
    }

    static void FormatShape(Shape shape, float width, float height, bool lockAspectRatio, float outlineWidth, Color outlineColor, Color fillColor) {
        shape.LockAspectRatio = lockAspectRatio;
        shape.Width = width;
        shape.Height = height;
        shape.Fill.SetSolidFill(fillColor);
        shape.Outline.SetSolidFill(outlineColor);
        shape.Outline.Width = outlineWidth;
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}

enum ShapeActionType {
    Picture,
    Textbox,
    Figure,
    Group
}
