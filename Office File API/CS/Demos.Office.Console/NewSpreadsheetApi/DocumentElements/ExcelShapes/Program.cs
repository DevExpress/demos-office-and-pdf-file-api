using System.Drawing;
using DevExpress.Docs.Office;
using DevExpress.Docs.Spreadsheet;
using DevExpress.Drawing;

class Program {
    static readonly string ImageFilePath = Path.Combine("Documents", "Photo1.jpeg");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Shape_Document.xlsx";
    const ShapeActionType ShapeAction = ShapeActionType.Picture;
    const bool Embedded = true;
    const string ShapeText = "Sample Text";
    const float ShapeWidth = 450;
    const float ShapeHeight = 450;
    const float OutlineWidth = 4;
    const int NumberOfShapes = 3;
    static readonly ShapeType ShapeGeometry = ShapeType.Rectangle;
    static readonly Color OutlineColor = ColorTranslator.FromHtml("#9E67FF");
    static readonly Color FillColor = ColorTranslator.FromHtml("#FFFFFF");
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var imagePath = Path.Combine(AppContext.BaseDirectory, ImageFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var imageStream = File.OpenRead(imagePath);
            using var outputStream = ShapeAction switch {
                ShapeActionType.Picture => GenerateWorkbookWithImage(imageStream, Embedded, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Textbox => GenerateWorkbookWithTextbox(ShapeText, ShapeWidth, ShapeHeight, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Figure => GenerateWorkbookWithFigure(ShapeGeometry, ShapeWidth, ShapeHeight, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                ShapeActionType.Group => GenerateWorkbookWithShapeGroup(ShapeGeometry, NumberOfShapes, ShapeWidth, ShapeHeight, FillColor, OutlineColor, OutlineWidth, OutputFormat),
                _ => GenerateWorkbookWithImage(imageStream, Embedded, OutlineColor, OutlineWidth, OutputFormat)
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

    static Stream GenerateWorkbookWithImage(Stream imageStream, bool embedded, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Images";

        if(embedded) {
            var cellImage = new CellImage(DXImage.FromStream(imageStream), "Sample Image", true);
            worksheet.Cells["B2"].SetImage(cellImage);
        }
        else {
            var imageShape = new PictureShape(DXImage.FromStream(imageStream), ShapeAnchorType.TwoCells) {
                ShapeAnchor = new TwoCellAnchor(1, 1, 11, 4),
                AlternativeText = "Sample Image"
            };

            imageShape.OutlineStyle = CreateOutline(outlineColor, outlineWidth);

            worksheet.Shapes.Add(imageShape);
        }

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithTextbox(string text, float width, float height,
        Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "TextBox";

        var textBox = new Shape(ShapeAnchorType.TopLeftCell) {
            ShapeAnchor = CreateCellAnchor(1, 1, width, height),
            TextArea = new TextArea(text)
        };
        FormatShape(textBox, outlineWidth, outlineColor, fillColor);

        worksheet.Shapes.Add(textBox);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithFigure(ShapeType geometry, float width, float height,
        Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Figure";

        var figure = worksheet.Shapes.AddShape(geometry, ShapeAnchorType.TopLeftCell);
        figure.ShapeAnchor = CreateCellAnchor(1, 1, width, height);
        FormatShape(figure, outlineWidth, outlineColor, fillColor);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream GenerateWorkbookWithShapeGroup(ShapeType geometry, int numberOfShapes, float width, float height,
        Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var workbook = new Workbook();

        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Shape Group";

        var shapes = new ShapeBase[numberOfShapes];
        for(int i = 0; i < numberOfShapes; i++) {
            var shape = new Shape(ShapeAnchorType.Absolute) {
                Type = geometry,
                ShapeAnchor = new AbsoluteAnchor {
                    X = i * 150,
                    Y = i * 150,
                    Width = width,
                    Height = height
                }
            };
            FormatShape(shape, outlineWidth, outlineColor, fillColor);
            shapes[i] = shape;
        }

        worksheet.Shapes.AddShapeGroup(ShapeAnchorType.Absolute, shapes);

        return CreateOutputStream(workbook, outputFormat);
    }

    static CellAnchor CreateCellAnchor(int topLeftRowIndex, int topLeftColumnIndex, float width, float height) {
        return new CellAnchor {
            TopLeftRowIndex = topLeftRowIndex,
            TopLeftColumnIndex = topLeftColumnIndex,
            Width = width,
            Height = height
        };
    }

    static void FormatShape(Shape shape, float outlineWidth, Color outlineColor, Color fillColor) {
        shape.Fill = new SolidFill(fillColor);
        shape.OutlineStyle = CreateOutline(outlineColor, outlineWidth);
    }

    static OutlineStyle CreateOutline(Color color, float width) {
        return new OutlineStyle {
            Fill = new SolidFill(color),
            Width = width
        };
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }
}

enum ShapeActionType {
    Picture,
    Textbox,
    Figure,
    Group
}
