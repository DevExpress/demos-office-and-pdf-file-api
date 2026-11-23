#region Code

using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using DevExpress.Drawing;
using System.Drawing;
        #endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentElements {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordShapesController : WordControllerBase {
        [HttpPost("shapes/generate")]
        public async Task<IActionResult> GenerateDocument([FromForm] WordShapesModel model) {
            // Determine target format
            var outputFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = new MemoryStream();

            var defaultLockAspectRatio = false;
            float defaultWidthInInches = 1.5f, defaultHeightInInches = 1.5f;

            var fillColor = ColorTranslator.FromHtml(model.FillColor ?? string.Empty);
            var outlineColor = ColorTranslator.FromHtml(model.OutlineColor ?? string.Empty);
            var outlineWidth = model.OutlineWidth;

            // Generate document with shapes
            switch(model.ShapeType) {
                case ShapeActionType.Picture:
                default:
                    using(var imageFile = new FileStream(DemoDocumentPaths.WordShapeImage, FileMode.Open, FileAccess.Read)) {
                        var imageStream = new MemoryStream();
                        await imageFile.CopyToAsync(imageStream);
                        imageStream.Position = 0;

                        outputStream = GenerateDocumentWithImage(
                            imageStream,
                            model.ImageScale,
                            defaultLockAspectRatio,
                            outlineColor,
                            outlineWidth,
                            outputFormat);
                    }

                    break;
                case ShapeActionType.Textbox:
                    outputStream = GenerateDocumentWithTextbox(
                        model.ShapeText ?? string.Empty,
                        defaultWidthInInches,
                        defaultHeightInInches,
                        defaultLockAspectRatio,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Figure:
                    outputStream = GenerateDocumentWithFigure(
                        GetShapeType(model.ShapeGeometry),
                        defaultWidthInInches,
                        defaultHeightInInches,
                        defaultLockAspectRatio,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Group:
                    outputStream = GenerateDocumentWithShapeGroup(
                        GetShapeType(model.ShapeGeometry),
                        model.NumberOfShapes,
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
        const float DocumentUnitsPerInch = 300f;
        const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

        Stream GenerateDocumentWithImage(Stream imageStream, float scaleFactor, bool lockAspectRatio,
                                         Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var document = new WordDocument();

            var picture = new PictureShape(DXImage.FromStream(imageStream));
            document.Body.AddParagraph().InlineElements.Add(picture);

            picture.LockAspectRatio = lockAspectRatio;
            picture.Width = picture.OriginalSize.Width * scaleFactor;
            picture.Height = picture.OriginalSize.Height * scaleFactor;
            picture.Outline = new LineStyle {
                Fill = new SolidFill(outlineColor),
                Width = outlineWidth * DocumentUnitsPerPoint
            };

            // Create output stream and save document in target format
            return CreateOutputStream(document, outputFormat);
        }

        Stream GenerateDocumentWithTextbox(string text, float width, float height, bool lockAspectRatio,
                                           Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var document = new WordDocument();

            var shape = new Shape(ShapeType.Rectangle) {
                Content = new ShapeTextBoxContent()
            };
            shape.Content.AddParagraph(text);
            FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            document.Body.AddParagraph().InlineElements.Add(shape);

            // Create output stream and save document in target format
            return CreateOutputStream(document, outputFormat);
        }

        Stream GenerateDocumentWithFigure(ShapeType geometry, float width, float height, bool lockAspectRatio,
                                          Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var document = new WordDocument();

            var shape = new Shape(geometry) {
                Anchor = new FloatingAnchor { Wrapping = TextWrapType.InFrontOfText }
            };
            FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            document.Body.AddParagraph().InlineElements.Add(shape);

            // Create output stream and save document in target format
            return CreateOutputStream(document, outputFormat);
        }

        Stream GenerateDocumentWithShapeGroup(ShapeType geometry, int numberOfShapes, float width, float height,
                                              bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var document = new WordDocument();

            var group = new GroupShape();
            document.Body.AddParagraph().InlineElements.Add(group);
            for(var i = 0; i < numberOfShapes; i++) {
                var shape = new Shape(geometry);
                FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
                var offset = i * 0.5f * DocumentUnitsPerInch;
                group.Shapes.Add(shape, new RectangleF(offset, offset, shape.Width, shape.Height));
            }

            // Create output stream and save document in target format
            return CreateOutputStream(document, outputFormat);
        }

        void FormatShape(Shape shape, float widthInInches, float heightInInches, bool lockAspectRatio, float outlineWidthPoints, Color outlineColor, Color fillColor) {
            shape.LockAspectRatio = lockAspectRatio;
            shape.Width = widthInInches * DocumentUnitsPerInch;
            shape.Height = heightInInches * DocumentUnitsPerInch;
            shape.Outline = new LineStyle {
                Fill = new SolidFill(outlineColor),
                Width = outlineWidthPoints * DocumentUnitsPerPoint
            };
            shape.Fill = new SolidFill(fillColor);
        }

        Stream CreateOutputStream(WordDocument document, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            document.Save(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}