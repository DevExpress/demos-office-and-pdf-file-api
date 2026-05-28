#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Docs.Office;
using DevExpress.Office.Utils;
using DevExpress.Drawing;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentElements;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentElements {
    [ApiController]
    [Route("api/presentation/shapes")]
    public class PresentationShapesController : PresentationControllerBase {
        [HttpPost("process")]
        public async Task<IActionResult> ProcessShapes([FromForm] PresentationShapesModel model) {
            var outputFormat = PresentationFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = new MemoryStream();

            bool defaultLockAspectRatio = false;
            float defaultXInInches = 1f, defaultYInInches = 1f;
            float defaultWidthInInches = 1.5f, defaultHeightInInches = 1.5f;

            Color fillColor = ColorTranslator.FromHtml(model.FillColor ?? string.Empty);
            Color outlineColor = ColorTranslator.FromHtml(model.OutlineColor ?? string.Empty);
            float outlineWidth = model.OutlineWidth;

            // Generate presentation with shapes
            switch(model.ShapeActionType) {
                case ShapeActionType.Picture:
                default:
                    using(var imageFile = new FileStream(DemoDocumentPaths.PresentationShapeImage, FileMode.Open, FileAccess.Read)) {
                        var imageStream = new MemoryStream();
                        await imageFile.CopyToAsync(imageStream);
                        imageStream.Position = 0;
                        outputStream = AddImage(
                            imageStream,
                            defaultLockAspectRatio,
                            defaultXInInches,
                            defaultYInInches,
                            defaultWidthInInches,
                            defaultHeightInInches,
                            false, 0, 0, 0, 0,
                            outlineColor,
                            outlineWidth,
                            outputFormat);
                    }
                    break;
                case ShapeActionType.Textbox:
                    outputStream = AddTextbox(
                        model.ShapeText ?? string.Empty,
                        defaultXInInches,
                        defaultYInInches,
                        defaultWidthInInches,
                        defaultHeightInInches,
                        defaultLockAspectRatio,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Figure:
                    outputStream = AddFigure(
                        GetShapeType(model.ShapeGeometryPresetType),
                        defaultXInInches,
                        defaultYInInches,
                        defaultWidthInInches,
                        defaultHeightInInches,
                        defaultLockAspectRatio,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Group:
                    outputStream = AddShapeGroup(
                        GetShapeType(model.ShapeGeometryPresetType),
                        model.NumberOfShapes,
                        defaultXInInches,
                        defaultYInInches,
                        defaultWidthInInches,
                        defaultHeightInInches,
                        defaultLockAspectRatio,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        ShapeType GetShapeType(GeometryPresetType geometryPresetType) {
            return geometryPresetType switch {
                GeometryPresetType.Rectangle => ShapeType.Rectangle,
                GeometryPresetType.Oval => ShapeType.Oval,
                GeometryPresetType.Star4 => ShapeType.Star4,
                GeometryPresetType.Heart => ShapeType.Heart,
                GeometryPresetType.RightArrow => ShapeType.RightArrow,
                _ => ShapeType.Rectangle
            };
        }

        #region Code
        Stream AddImage(Stream imageStream, bool lockAspectRatio,
            float x, float y, float width, float height,
            bool shouldCrop, float cropLeft, float cropRight, float cropTop, float cropBottom,
            Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var presentation = new Presentation();
            presentation.Slides.Clear();

            var slide = new Slide(SlideLayoutType.Blank);
            presentation.Slides.Add(slide);

            var pictureShape = new PictureShape() {
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

            // Create output stream and save document in target format
            return CreateOutputStream(presentation, outputFormat);
        }

        Stream AddTextbox(string text, float x, float y, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var presentation = new Presentation();
            presentation.Slides.Clear();

            var slide = new Slide(SlideLayoutType.Blank);
            presentation.Slides.Add(slide);

            var textboxShape = new Shape(ShapeType.Rectangle) {
                TextArea = new TextArea(text),
            };

            textboxShape.TextArea.Level1ParagraphProperties.TextProperties.Fill = new SolidFill(Color.Black);
            FormatShape(textboxShape, x, y, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            slide.Shapes.Add(textboxShape);

            // Create output stream and save document in target format
            return CreateOutputStream(presentation, outputFormat);
        }

        Stream AddFigure(ShapeType shapeType, float x, float y, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var presentation = new Presentation();
            presentation.Slides.Clear();

            var slide = new Slide(SlideLayoutType.Blank);
            presentation.Slides.Add(slide);

            var shape = new Shape(shapeType);
            FormatShape(shape, x, y, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            slide.Shapes.Add(shape);

            // Create output stream and save document in target format
            return CreateOutputStream(presentation, outputFormat);
        }

        Stream AddShapeGroup(ShapeType shapeType, int numberOfShapes,
            float x, float y, float width, float height,
            bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
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

            // Create output stream and save document in target format
            return CreateOutputStream(presentation, outputFormat);
        }

        void FormatShape(Shape shape, float x, float y, float width, float height, bool lockAspectRatio, float outlineWidth, Color outlineColor, Color fillColor) {
            shape.X = Units.InchesToDocumentsF(x);
            shape.Y = Units.InchesToDocumentsF(y);
            shape.Width = Units.InchesToDocumentsF(width);
            shape.Height = Units.InchesToDocumentsF(height);
            shape.Fill = new SolidFill(fillColor);
            shape.Outline = new LineStyle { Fill = new SolidFill(outlineColor), Width = Units.PointsToDocumentsF(outlineWidth) };
            shape.LockSettings.DisableAspectRatioChange = lockAspectRatio;
        }

        Stream CreateOutputStream(Presentation presentation, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            presentation.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
