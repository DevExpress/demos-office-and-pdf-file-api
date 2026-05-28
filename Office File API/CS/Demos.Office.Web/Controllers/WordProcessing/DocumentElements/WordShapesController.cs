#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/shapes")]
    public class WordShapesController : WordControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateDocument([FromForm] WordShapesModel model) {
            // Determine target format
            var outputFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = new MemoryStream();

            bool defaultLockAspectRatio = false;
            float defaultWidthInInches = 1.5f, defaultHeightInInches = 1.5f;

            Color fillColor = ColorTranslator.FromHtml(model.FillColor ?? string.Empty);
            Color outlineColor = ColorTranslator.FromHtml(model.OutlineColor ?? string.Empty);
            float outlineWidth = model.OutlineWidth;

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
                            false, 0, 0, 0, 0,
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
                        model.ShapeGeometry,
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
                        model.ShapeGeometry,
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

        #region Code
        Stream GenerateDocumentWithImage(Stream imageStream, float scaleFactor, bool lockAspectRatio,
            bool shouldCrop, float cropLeft, float cropRight, float cropTop, float cropBottom,
            Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            var shape = document.Shapes.InsertPicture(document.Range.End, DocumentImageSource.FromStream(imageStream));
            shape.LockAspectRatio = lockAspectRatio;
            shape.TextWrapping = TextWrappingType.InLineWithText;
            shape.ScaleX = scaleFactor;
            shape.ScaleY = scaleFactor;

            if(shouldCrop) {
                RectangleOffset cropOffset = new RectangleOffset(cropLeft, cropTop, cropRight, cropBottom);
                shape.PictureFormat.SourceRect = cropOffset;
            }

            shape.Line.Color = outlineColor;
            shape.Line.Thickness = outlineWidth;

            // Create output stream and save document in target format
            return CreateOutputStream(wordProcessor, outputFormat);
        }

        Stream GenerateDocumentWithTextbox(string text, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            var shape = document.Shapes.InsertTextBox(document.Range.End);
            shape.ShapeFormat.TextBox.Document.AppendText(text);
            shape.TextWrapping = TextWrappingType.InLineWithText;
            FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);

            // Create output stream and save document in target format
            return CreateOutputStream(wordProcessor, outputFormat);
        }

        Stream GenerateDocumentWithFigure(ShapeGeometryPreset geometry, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            var shape = document.Shapes.InsertShape(document.Range.End, geometry);
            FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);

            // Create output stream and save document in target format
            return CreateOutputStream(wordProcessor, outputFormat);
        }

        Stream GenerateDocumentWithShapeGroup(ShapeGeometryPreset geometry, int numberOfShapes, float width, float height,
            bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            // Set document unit to inches for precise sizing
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            var group = document.Shapes.InsertGroup(document.Range.End);
            group.TextWrapping = TextWrappingType.InLineWithText;
            var groupItems = group.GroupItems;
            for(int i = 0; i < numberOfShapes; i++) {
                var shape = groupItems.AddShape(geometry, new RectangleF(i * 0.5f, i * 0.5f, width, height));
                FormatShape(shape, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            }

            // Create output stream and save document in target format
            return CreateOutputStream(wordProcessor, outputFormat);
        }

        void FormatShape(DrawingObject shape, float width, float height, bool lockAspectRatio, float outlineWidthPoints, Color outlineColor, Color fillColor) {
            shape.LockAspectRatio = lockAspectRatio;
            shape.Width = width;
            shape.Height = height;
            shape.Line.Color = outlineColor;
            shape.Line.Thickness = outlineWidthPoints;
            shape.Fill.Color = fillColor;
        }

        Stream CreateOutputStream(RichEditDocumentServer wordProcessor, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            wordProcessor.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
