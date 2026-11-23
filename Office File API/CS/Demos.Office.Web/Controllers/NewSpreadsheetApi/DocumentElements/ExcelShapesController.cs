#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Docs.Office;
using DevExpress.Drawing;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentElements {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelShapesController : SpreadsheetControllerBase {
        [HttpPost("shapes/generate")]
        public async Task<IActionResult> Generate([FromForm] ExcelShapesModel model) {
            var outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = new MemoryStream();

            float defaultWidth = 450, defaultHeight = 450;

            Color fillColor = ColorTranslator.FromHtml(model.FillColor ?? string.Empty);
            Color outlineColor = ColorTranslator.FromHtml(model.OutlineColor ?? string.Empty);
            float outlineWidth = model.OutlineWidth;

            switch(model.ShapeType) {
                case ShapeActionType.Picture:
                default:
                    using(var imageFile = new FileStream(DemoDocumentPaths.SpreadsheetShapeImage, FileMode.Open, FileAccess.Read)) {
                        var imageStream = new MemoryStream();
                        await imageFile.CopyToAsync(imageStream);
                        imageStream.Position = 0;

                        outputStream = GenerateWorkbookWithImage(
                            imageStream,
                            model.Embedded,
                            outlineColor,
                            outlineWidth,
                            outputFormat);
                    }
                    break;
                case ShapeActionType.Textbox:
                    outputStream = GenerateWorkbookWithTextbox(
                        model.ShapeText ?? string.Empty,
                        defaultWidth,
                        defaultHeight,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Figure:
                    outputStream = GenerateWorkbookWithFigure(
                        GetShapeType(model.ShapeGeometry),
                        defaultWidth,
                        defaultHeight,
                        fillColor,
                        outlineColor,
                        outlineWidth,
                        outputFormat);
                    break;
                case ShapeActionType.Group:
                    outputStream = GenerateWorkbookWithShapeGroup(
                        GetShapeType(model.ShapeGeometry),
                        model.NumberOfShapes,
                        defaultWidth,
                        defaultHeight,
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
        Stream GenerateWorkbookWithImage(Stream imageStream, bool embedded, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
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

                imageShape.Outline = CreateOutline(outlineColor, outlineWidth);

                worksheet.Shapes.Add(imageShape);
            }

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream GenerateWorkbookWithTextbox(string text, float width, float height,
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

        Stream GenerateWorkbookWithFigure(ShapeType geometry, float width, float height,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Figure";

            var figure = worksheet.Shapes.AddShape(geometry, ShapeAnchorType.TopLeftCell);
            figure.ShapeAnchor = CreateCellAnchor(1, 1, width, height);
            FormatShape(figure, outlineWidth, outlineColor, fillColor);

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream GenerateWorkbookWithShapeGroup(ShapeType geometry, int numberOfShapes, float width, float height,
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

        CellAnchor CreateCellAnchor(int topLeftRowIndex, int topLeftColumnIndex, float width, float height) {
            return new CellAnchor {
                TopLeftRowIndex = topLeftRowIndex,
                TopLeftColumnIndex = topLeftColumnIndex,
                Width = width,
                Height = height
            };
        }

        void FormatShape(Shape shape, float outlineWidth, Color outlineColor, Color fillColor) {
            shape.Fill = new SolidFill(fillColor);
            shape.Outline = CreateOutline(outlineColor, outlineWidth);
        }

        LineStyle CreateOutline(Color color, float width) {
            return new LineStyle {
                Fill = new SolidFill(color),
                Width = width
            };
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
