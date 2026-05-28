using BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements;
using Microsoft.AspNetCore.Mvc;
using System.Drawing;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.DocumentElements {
    [ApiController]
    [Route("api/wordprocessing/tables")]
    public class WordTablesController : WordControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateTableDocument([FromForm] WordTablesModel model) {

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with table
            var outputStream = GenerateDocumentWithTable(model, targetFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);

        }

        Stream GenerateDocumentWithTable(WordTablesModel model, DocumentFormat targetFormat) {
            using var wordProcessor = new RichEditDocumentServer();
            var document = wordProcessor.Document;

            var table = GenerateTable(document, model.NumberOfRows, model.NumberOfColumns, model.TableSizeMode, model.FixedColumnWidth);

            // Apply border formatting
            ApplyTableBorders(document, table, model.BorderType, model.BorderStyle, ColorTranslator.FromHtml(model.BorderColor ?? string.Empty), model.BorderWidth);

            // Apply text wrapping and positioning
            ApplyTableWrappingAndPosition(document, table, model.TextWrapping, model.HorizontalPosition, model.VerticalPosition);

            // Populate table with data
            PopulateTableData(document, table, model.IncludeHeaders, model.IncludeSampleData);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, targetFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #region Code
        Table GenerateTable(Document document, int rowCount, int colCount, AutoFitBehaviorType tableAutoFit, float colWidth) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            Table table = document.Tables.Create(document.Range.End, rowCount, colCount, tableAutoFit);

            if(tableAutoFit == AutoFitBehaviorType.FixedColumnWidth) {

                table.ForEachCell((cell, rowIndex, colIndex) => {
                    cell.PreferredWidthType = WidthType.Fixed;
                    cell.PreferredWidth = colWidth;
                });
            }

            return table;
        }

        void ApplyTableBorders(Document document, Table table, BorderType borderType,
            BorderLineStyle borderStyle, Color borderColor, float borderThickness) {
            switch(borderType) {
                case BorderType.NoBorders:
                    ClearTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right,
                        table.Borders.InsideHorizontalBorder, table.Borders.InsideVerticalBorder });
                    break;
                case BorderType.AllBorders:
                    ApplyTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right,
                        table.Borders.InsideHorizontalBorder, table.Borders.InsideVerticalBorder },
                        borderStyle, borderColor, borderThickness);
                    break;
                case BorderType.OutsideBorders:
                    ApplyTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right },
                        borderStyle, borderColor, borderThickness);
                    ClearTableBorders(new List<TableBorder> {
                        table.Borders.InsideHorizontalBorder, table.Borders.InsideVerticalBorder });
                    break;
                case BorderType.InsideBorders:
                    ApplyTableBorders(new List<TableBorder> {
                        table.Borders.InsideHorizontalBorder, table.Borders.InsideVerticalBorder },
                        borderStyle, borderColor, borderThickness);
                    ClearTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right });
                    break;
                case BorderType.InsideHorizontalBorders:
                    ApplyTableBorders(new List<TableBorder> {
                        table.Borders.InsideHorizontalBorder },
                        borderStyle, borderColor, borderThickness);
                    ClearTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right,
                        table.Borders.InsideVerticalBorder });
                    break;
                case BorderType.InsideVerticalBorders:
                    ApplyTableBorders(new List<TableBorder> {
                        table.Borders.InsideVerticalBorder },
                        borderStyle, borderColor, borderThickness);
                    ClearTableBorders(new List<TableBorder> {
                        table.Borders.Top, table.Borders.Bottom, table.Borders.Left, table.Borders.Right,
                        table.Borders.InsideHorizontalBorder });
                    break;
            }
        }

        void ApplyTableBorders(List<TableBorder> tableBorders, BorderLineStyle style, Color color, float thickness) {
            foreach(var border in tableBorders) {
                border.LineStyle = style;
                border.LineColor = color;
                border.LineThickness = thickness;
            }
        }

        void ClearTableBorders(List<TableBorder> tableBorders) {
            foreach(var border in tableBorders)
                border.LineStyle = BorderLineStyle.None;
        }

        void ApplyTableWrappingAndPosition(Document document, Table table, TableTextWrappingType textWrappingType, float positionX, float positionY) {
            document.Unit = DevExpress.Office.DocumentUnit.Inch;

            table.TextWrappingType = textWrappingType;

            if(textWrappingType == TableTextWrappingType.Around) {
                table.RelativeHorizontalPosition = TableRelativeHorizontalPosition.Page;
                table.OffsetXRelative = positionX;

                table.RelativeVerticalPosition = TableRelativeVerticalPosition.Page;
                table.OffsetYRelative = positionY;
            }
        }

        void PopulateTableData(Document document, Table table, bool includeHeaders, bool includeSampleData) {
            if(includeHeaders) {
                TableRow headerRow = table.FirstRow;
                for(int col = 0; col < table.Rows[0].Cells.Count; col++) {
                    var headerCell = headerRow.Cells[col];
                    document.InsertText(headerCell.Range.Start, $"Column {col + 1}");

                    headerCell.BackgroundColor = Color.LightBlue;
                }
                // Format header cells
                var headerCharProps = document.BeginUpdateCharacters(headerRow.Range);
                headerCharProps.Bold = true;
                headerCharProps.ForeColor = Color.White;
                document.EndUpdateCharacters(headerCharProps);

                var headerParaProps = document.BeginUpdateParagraphs(headerRow.Range);
                headerParaProps.Alignment = ParagraphAlignment.Center;
                document.EndUpdateParagraphs(headerParaProps);
            }

            // Populate data rows
            if(includeSampleData) {
                int startRow = includeHeaders ? 1 : 0;
                for(int row = startRow; row < table.Rows.Count; row++) {
                    for(int col = 0; col < table.Rows[row].Cells.Count; col++) {
                        var cell = table.Rows[row].Cells[col];

                        // Add sample data
                        var cellText = col == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{col + 1}";
                        document.InsertText(cell.Range.Start, cellText);
                        // Set vertical alignment
                        cell.VerticalAlignment = TableCellVerticalAlignment.Center;
                    }
                    // Center align cell content
                    var cellParaProps = document.BeginUpdateParagraphs(table.Rows[row].Range);
                    cellParaProps.Alignment = ParagraphAlignment.Center;
                    document.EndUpdateParagraphs(cellParaProps);
                }
            }
        }
        #endregion Code
    }
}
