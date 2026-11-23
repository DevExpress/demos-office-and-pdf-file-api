using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;

#region Code

using DevExpress.Docs;
using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.DocumentElements {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordTablesController : WordControllerBase {
        [HttpPost("tables/generate")]
        public IActionResult GenerateTableDocument([FromForm] WordTablesModel model) {
            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with table
            var outputStream = GenerateDocumentWithTable(model, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        Stream GenerateDocumentWithTable(WordTablesModel model, DocumentFormat targetFormat) {
            using var document = new WordDocument();

            var table = GenerateTable(model.NumberOfRows, model.NumberOfColumns, model.TableSizeMode, model.FixedColumnWidth);
            // Apply border formatting
            ApplyTableBorders(table, model.BorderType, model.BorderStyle, ColorTranslator.FromHtml(model.BorderColor ?? string.Empty), model.BorderWidth);
            // Apply text wrapping and positioning
            ApplyTableWrappingAndPosition(table, model.TextWrapping, model.HorizontalPosition, model.VerticalPosition);
            // Populate table with data
            PopulateTableData(table, model.IncludeHeaders, model.IncludeSampleData);
            document.Body.Add(table);

            var outputStream = new MemoryStream();
            document.Save(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        const float DocumentUnitsPerInch = 300f;
        const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

        Table GenerateTable(int rowCount, int colCount, TableSizeModeType tableSizeMode, float colWidth) {
            var table = new Table(rowCount, colCount);

            switch(tableSizeMode) {
                case TableSizeModeType.AutoFitToWindow:
                    table.Layout = TableLayoutType.AutoFit;
                    table.PreferredWidth = TableWidth.Percent(100);
                    break;
                case TableSizeModeType.AutoFitToContents:
                    table.Layout = TableLayoutType.AutoFit;
                    table.PreferredWidth = TableWidth.Auto;
                    break;
                case TableSizeModeType.FixedColumnWidth:
                    table.Layout = TableLayoutType.Fixed;
                    foreach(var row in table.Rows) {
                        foreach(var cell in row.Cells)
                            cell.PreferredWidth = TableWidth.Fixed(colWidth * DocumentUnitsPerInch);
                    }

                    break;
            }

            return table;
        }

        void ApplyTableBorders(Table table, TableBorderType borderType,
                               BorderLineStyle borderStyle, Color borderColor, float borderThickness) {
            var visibleBorder = new TableBorder(borderStyle, borderThickness * DocumentUnitsPerPoint, new OfficeColor(borderColor));
            var hiddenBorder = new TableBorder(BorderLineStyle.None);

            var hasOutsideBorders = borderType is TableBorderType.AllBorders or TableBorderType.OutsideBorders;
            var hasInsideHorizontalBorder = borderType is TableBorderType.AllBorders or TableBorderType.InsideBorders
                                                                                     or TableBorderType.InsideHorizontalBorders;
            var hasInsideVerticalBorder = borderType is TableBorderType.AllBorders or TableBorderType.InsideBorders
                                                                                   or TableBorderType.InsideVerticalBorders;

            var borders = table.Properties.Borders;
            borders.Top = hasOutsideBorders ? visibleBorder : hiddenBorder;
            borders.Bottom = hasOutsideBorders ? visibleBorder : hiddenBorder;
            borders.Left = hasOutsideBorders ? visibleBorder : hiddenBorder;
            borders.Right = hasOutsideBorders ? visibleBorder : hiddenBorder;
            borders.InsideHorizontal = hasInsideHorizontalBorder ? visibleBorder : hiddenBorder;
            borders.InsideVertical = hasInsideVerticalBorder ? visibleBorder : hiddenBorder;
        }

        void ApplyTableWrappingAndPosition(Table table, TableTextWrapping textWrapping, float positionX, float positionY) {
            table.Positioning.TextWrapping = textWrapping;

            if(textWrapping == TableTextWrapping.Around) {
                table.Positioning.HorizontalPosition = new TableOffsetHorizontalPosition(positionX * DocumentUnitsPerInch, TableRelativeHorizontalPosition.Page);
                table.Positioning.VerticalPosition = new TableOffsetVerticalPosition(positionY * DocumentUnitsPerInch, TableRelativeVerticalPosition.Page);
            }
        }

        void PopulateTableData(Table table, bool includeHeaders, bool includeSampleData) {
            if(includeHeaders) {
                var headerRow = table.Rows[0];
                for(var col = 0; col < headerRow.Cells.Count; col++) {
                    var headerCell = headerRow.Cells[col];
                    var headerText = new TextFragment($"Column {col + 1}");
                    // Format header cells
                    headerText.CharacterProperties.Bold = true;
                    headerText.CharacterProperties.ForeColor = new OfficeColor(Color.White);
                    AddCenteredParagraph(headerCell, headerText);

                    headerCell.Properties.BackColor = new OfficeColor(Color.LightBlue);
                }
            }

            // Populate data rows
            if(includeSampleData) {
                var startRow = includeHeaders ? 1 : 0;
                for(var row = startRow; row < table.Rows.Count; row++) {
                    for(var col = 0; col < table.Rows[row].Cells.Count; col++) {
                        var cell = table.Rows[row].Cells[col];

                        // Add sample data
                        var cellText = col == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{col + 1}";
                        AddCenteredParagraph(cell, new TextFragment(cellText));
                        // Set vertical alignment
                        cell.Properties.VerticalAlignment = TableCellVerticalAlignment.Center;
                    }
                }
            }
        }

        void AddCenteredParagraph(TableCell cell, TextFragment text) {
            var paragraph = new Paragraph();
            paragraph.InlineElements.Add(text);
            // Center align cell content
            paragraph.Properties.Alignment = ParagraphAlignment.Center;
            cell.Content.Add(paragraph);
        }
        #endregion Code
    }
}