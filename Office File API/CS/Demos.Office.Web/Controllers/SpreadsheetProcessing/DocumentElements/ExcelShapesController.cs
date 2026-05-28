#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelShapesController : SpreadsheetControllerBase {
        [HttpPost("shapes/generate")]
        public async Task<IActionResult> Generate([FromForm] ExcelShapesModel model) {
            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);
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
                    using(var imageFile = new FileStream(DemoDocumentPaths.SpreadsheetShapeImage, FileMode.Open, FileAccess.Read)) {
                        var imageStream = new MemoryStream();
                        await imageFile.CopyToAsync(imageStream);
                        imageStream.Position = 0;

                        outputStream = GenerateWorkbookWithImage(
                            imageStream,
                            model.Embedded,
                            defaultLockAspectRatio,
                            outlineColor,
                            outlineWidth,
                            outputFormat);
                    }
                    break;
                case ShapeActionType.Textbox:
                    outputStream = GenerateWorkbookWithTextbox(
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
                    outputStream = GenerateWorkbookWithFigure(
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
                    outputStream = GenerateWorkbookWithShapeGroup(
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
        Stream GenerateWorkbookWithImage(Stream imageStream, bool embedded, bool lockAspectRatio,
            Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Images";

            if(embedded) {
                var cell = worksheet.Cells["B2"];
                // Insert cell image from a stream
                cell.Value = imageStream;

                // Specify image information
                if(cell.Value.IsCellImage) {
                    cell.ImageInfo.Decorative = true;
                    cell.ImageInfo.AlternativeText = "Sample Image";
                }
            }
            else {
                var imageSource = SpreadsheetImageSource.FromStream(imageStream);
                var imageRange = worksheet.Range["B2:D11"];

                // Insert an image to fit in the specified range
                var imageShape = worksheet.Pictures.AddPicture(imageSource, imageRange);

                // Specify image settings
                imageShape.LockAspectRatio = lockAspectRatio;
                imageShape.Outline.SetSolidFill(outlineColor);
                imageShape.Outline.Width = outlineWidth;
            }

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream GenerateWorkbookWithTextbox(string text, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "TextBox";

            var shapeRange = worksheet.Range["B2:D3"];

            // Insert a textbox shape
            var textBox = worksheet.Shapes.AddTextBox(shapeRange, text);
            FormatShape(textBox, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream GenerateWorkbookWithFigure(ShapeGeometryPreset geometry, float width, float height, bool lockAspectRatio,
            Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Figure";

            var shapeRange = worksheet.Range["B2:D3"];

            // Insert a figure shape
            var figure = worksheet.Shapes.AddShape(geometry, shapeRange);
            FormatShape(figure, width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream GenerateWorkbookWithShapeGroup(ShapeGeometryPreset geometry, int numberOfShapes, float width, float height,
            bool lockAspectRatio, Color fillColor, Color outlineColor, float outlineWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch;

            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Shape Group";

            // Create shapes
            Shape[] shapes = new Shape[numberOfShapes];
            for(int i = 0; i < numberOfShapes; i++) {
                shapes[i] = worksheet.Shapes.AddShape(geometry, i * 0.5f, i * 0.5f, width, height);
                FormatShape(shapes[i], width, height, lockAspectRatio, outlineWidth, outlineColor, fillColor);
            }

            // Combine shapes into a group 
            Shape shapeGroup = worksheet.Shapes.GroupShapes(shapes);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        void FormatShape(Shape shape, float width, float height, bool lockAspectRatio, float outlineWidth, Color outlineColor, Color fillColor) {
            shape.LockAspectRatio = lockAspectRatio;
            shape.Width = width;
            shape.Height = height;
            shape.Fill.SetSolidFill(fillColor);
            shape.Outline.SetSolidFill(outlineColor);
            shape.Outline.Width = outlineWidth;
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);
            return outputStream;
        }
        #endregion Code
    }
}
