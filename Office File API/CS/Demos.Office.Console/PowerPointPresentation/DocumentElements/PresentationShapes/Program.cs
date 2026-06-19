using System.Drawing;
using DevExpress.Docs.Office;
using DevExpress.Docs.Presentation;
using DevExpress.Drawing;
using DevExpress.Office.Utils;

class Program {
    static readonly string ImageFilePath = Path.Combine("Documents", "Photo1.jpeg");
    const string OutputFileName = "Presentation_Shapes.pptx";
    const string OutputFormat = "pptx";
    const ShapeActionType ShapeAction = ShapeActionType.Picture;
    const GeometryPresetType ShapeGeometry = GeometryPresetType.Rectangle;
    const string ShapeText = "Sample Text";
    const float OutlineWidth = 1.0f;
    const string OutlineColor = "#9E67FF";
    const string FillColor = "#FFFFFF";
    const int NumberOfShapes = 3;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            var outputFormat = GetDocumentFormat(OutputFormat);
            Color fillColor = ColorTranslator.FromHtml(FillColor);
            Color outlineColor = ColorTranslator.FromHtml(OutlineColor);
            const bool defaultLockAspectRatio = false;
            const float defaultXInInches = 1f;
            const float defaultYInInches = 1f;
            const float defaultWidthInInches = 1.5f;
            const float defaultHeightInInches = 1.5f;

            using Stream outputStream = ShapeAction switch {
                ShapeActionType.Textbox => AddTextbox(ShapeText, defaultXInInches, defaultYInInches, defaultWidthInInches, defaultHeightInInches, defaultLockAspectRatio, fillColor, outlineColor, OutlineWidth, outputFormat),
                ShapeActionType.Figure => AddFigure(GetShapeType(ShapeGeometry), defaultXInInches, defaultYInInches, defaultWidthInInches, defaultHeightInInches, defaultLockAspectRatio, fillColor, outlineColor, OutlineWidth, outputFormat),
                ShapeActionType.Group => AddShapeGroup(GetShapeType(ShapeGeometry), NumberOfShapes, defaultXInInches, defaultYInInches, defaultWidthInInches, defaultHeightInInches, defaultLockAspectRatio, fillColor, outlineColor, OutlineWidth, outputFormat),
                _ => AddPicture(defaultLockAspectRatio, defaultXInInches, defaultYInInches, defaultWidthInInches, defaultHeightInInches, outlineColor, OutlineWidth, outputFormat)
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

    static ShapeType GetShapeType(GeometryPresetType geometryPresetType) {
        return geometryPresetType switch {
            GeometryPresetType.Rectangle => ShapeType.Rectangle,
            GeometryPresetType.Oval => ShapeType.Oval,
            GeometryPresetType.Star4 => ShapeType.Star4,
            GeometryPresetType.Heart => ShapeType.Heart,
            GeometryPresetType.RightArrow => ShapeType.RightArrow,
            _ => ShapeType.Rectangle
        };
    }

    static Stream AddPicture(bool lockAspectRatio, float x, float y, float width, float height, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        var imagePath = Path.Combine(AppContext.BaseDirectory, ImageFilePath);
        using var imageStream = File.OpenRead(imagePath);
        return AddImage(imageStream, lockAspectRatio, x, y, width, height, false, 0, 0, 0, 0, outlineColor, outlineWidth, outputFormat);
    }

    static Stream AddImage(Stream imageStream, bool lockAspectRatio, float x, float y, float width, float height, bool shouldCrop, float cropLeft, float cropRight, float cropTop, float cropBottom, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var presentation = new Presentation();
        presentation.Slides.Clear();

        var slide = new Slide(SlideLayoutType.Blank);
        presentation.Slides.Add(slide);

        var pictureShape = new PictureShape {
            Image = DXImage.FromStream(imageStream),
            X = Units.InchesToDocumentsF(x),
            Y = Units.InchesToDocumentsF(y),
            Width = Units.InchesToDocumentsF(width),
            Height = Units.InchesToDocumentsF(height),
            Outline = new LineStyle { Fill = new SolidFill(outlineColor), Width = Units.PointsToDocumentsF(outlineWidth) }
        };

        if(shouldCrop) {
            pictureShape.CropLeft = cropLeft;
            pictureShape.CropTop = cropTop;
            pictureShape.CropRight = cropRight;
            pictureShape.CropBottom = cropBottom;
        }

        pictureShape.LockSettings.DisableAspectRatioChange = lockAspectRatio;
        slide.Shapes.Add(pictureShape);
        return CreateOutputStream(presentation, outputFormat);
    }

    static Stream AddTextbox(string text, float x, float y, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var presentation = new Presentation();
        presentation.Slides.Clear();

        var slide = new Slide(SlideLayoutType.Blank);
        presentation.Slides.Add(slide);

        var textboxShape = new Shape(ShapeType.Rectangle) {
            TextArea = new TextArea(text)
        };

        textboxShape.TextArea.Level1ParagraphProperties.TextProperties.Fill = new SolidFill(Color.Black);
        FormatShape(textboxShape, x, y, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
        slide.Shapes.Add(textboxShape);
        return CreateOutputStream(presentation, outputFormat);
    }

    static Stream AddFigure(ShapeType shapeType, float x, float y, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var presentation = new Presentation();
        presentation.Slides.Clear();

        var slide = new Slide(SlideLayoutType.Blank);
        presentation.Slides.Add(slide);

        var shape = new Shape(shapeType);
        FormatShape(shape, x, y, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
        slide.Shapes.Add(shape);
        return CreateOutputStream(presentation, outputFormat);
    }

    static Stream AddShapeGroup(ShapeType shapeType, int numberOfShapes, float x, float y, float width, float height, bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
        using var presentation = new Presentation();
        presentation.Slides.Clear();

        var slide = new Slide(SlideLayoutType.Blank);
        presentation.Slides.Add(slide);

        var groupShape = new GroupShape();
        for(int i = 0; i < numberOfShapes; i++) {
            var shape = new Shape(shapeType);
            FormatShape(shape, x, y, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            groupShape.Shapes.Add(shape);
            x += 0.5f;
            y += 0.5f;
        }

        slide.Shapes.Add(groupShape);
        return CreateOutputStream(presentation, outputFormat);
    }

    static void FormatShape(Shape shape, float x, float y, float width, float height, bool lockAspectRatio, float outlineWidth, Color outlineColor, Color fillColor) {
        shape.X = Units.InchesToDocumentsF(x);
        shape.Y = Units.InchesToDocumentsF(y);
        shape.Width = Units.InchesToDocumentsF(width);
        shape.Height = Units.InchesToDocumentsF(height);
        shape.Fill = new SolidFill(fillColor);
        shape.Outline = new LineStyle { Fill = new SolidFill(outlineColor), Width = Units.PointsToDocumentsF(outlineWidth) };
        shape.LockSettings.DisableAspectRatioChange = lockAspectRatio;
    }

    static Stream CreateOutputStream(Presentation presentation, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        presentation.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "pptm" => DocumentFormat.Pptm,
            "potx" => DocumentFormat.Potx,
            "potm" => DocumentFormat.Potm,
            _ => DocumentFormat.Pptx
        };
    }

    enum ShapeActionType {
        Picture,
        Textbox,
        Figure,
        Group
    }
}
