using DevExpress.Docs;
using DevExpress.Docs.Word;
using System.Drawing;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Table_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const int NumberOfRows = 5;
    const int NumberOfColumns = 4;
    const TableSizeModeType TableSizeMode = TableSizeModeType.AutoFitToWindow;
    const float FixedColumnWidth = 1.5f;
    const TableBorderType BorderType = TableBorderType.AllBorders;
    const BorderLineStyle BorderStyle = BorderLineStyle.Single;
    static readonly Color BorderColor = ColorTranslator.FromHtml("#000000");
    const float BorderWidth = 1f;
    const TableTextWrapping TextWrapping = TableTextWrapping.None;
    const float HorizontalPosition = 0f;
    const float VerticalPosition = 0f;
    const bool IncludeHeaders = true;
    const bool IncludeSampleData = true;

    const float DocumentUnitsPerInch = 300f;
    const float DocumentUnitsPerPoint = DocumentUnitsPerInch / 72f;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var outputStream = GenerateDocumentWithTable(OutputFormat);
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

    static Stream GenerateDocumentWithTable(DocumentFormat targetFormat) {
        using var document = new WordDocument();

        var table = GenerateTable(NumberOfRows, NumberOfColumns, TableSizeMode, FixedColumnWidth);
        ApplyTableBorders(table, BorderType, BorderStyle, BorderColor, BorderWidth);
        ApplyTableWrappingAndPosition(table, TextWrapping, HorizontalPosition, VerticalPosition);
        PopulateTableData(table, IncludeHeaders, IncludeSampleData);
        document.Body.Add(table);

        var outputStream = new MemoryStream();
        document.Save(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static Table GenerateTable(int rowCount, int colCount, TableSizeModeType tableSizeMode, float colWidth) {
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

    static void ApplyTableBorders(Table table, TableBorderType borderType,
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

    static void ApplyTableWrappingAndPosition(Table table, TableTextWrapping textWrapping, float positionX, float positionY) {
        table.Positioning.TextWrapping = textWrapping;

        if(textWrapping == TableTextWrapping.Around) {
            table.Positioning.HorizontalPosition = new TableOffsetHorizontalPosition(positionX * DocumentUnitsPerInch, TableRelativeHorizontalPosition.Page);
            table.Positioning.VerticalPosition = new TableOffsetVerticalPosition(positionY * DocumentUnitsPerInch, TableRelativeVerticalPosition.Page);
        }
    }

    static void PopulateTableData(Table table, bool includeHeaders, bool includeSampleData) {
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

    static void AddCenteredParagraph(TableCell cell, TextFragment text) {
        var paragraph = new Paragraph();
        paragraph.InlineElements.Add(text);
        // Center align cell content
        paragraph.Properties.Alignment = ParagraphAlignment.Center;
        cell.Content.Add(paragraph);
    }
}

enum TableSizeModeType {
    AutoFitToWindow,
    AutoFitToContents,
    FixedColumnWidth
}

enum TableBorderType {
    NoBorders,
    AllBorders,
    OutsideBorders,
    InsideBorders,
    InsideHorizontalBorders,
    InsideVerticalBorders
}