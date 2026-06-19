using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    const string OutputFileName = "Table_Result.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const int NumberOfRows = 5;
    const int NumberOfColumns = 4;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
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
        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;
        document.Unit = DevExpress.Office.DocumentUnit.Inch;

        var table = document.Tables.Create(document.Range.End, NumberOfRows, NumberOfColumns, AutoFitBehaviorType.AutoFitToWindow);
        ApplyBorders(table, BorderLineStyle.Single, Color.Black, 1f);
        PopulateTableData(document, table);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, targetFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static void ApplyBorders(Table table, BorderLineStyle style, Color color, float thickness) {
        foreach(var border in new[] {
            table.Borders.Top,
            table.Borders.Bottom,
            table.Borders.Left,
            table.Borders.Right,
            table.Borders.InsideHorizontalBorder,
            table.Borders.InsideVerticalBorder
        }) {
            border.LineStyle = style;
            border.LineColor = color;
            border.LineThickness = thickness;
        }
    }

    static void PopulateTableData(Document document, Table table) {
        var headerRow = table.FirstRow;
        for(int col = 0; col < table.Rows[0].Cells.Count; col++) {
            var headerCell = headerRow.Cells[col];
            document.InsertText(headerCell.Range.Start, $"Column {col + 1}");
            headerCell.BackgroundColor = Color.LightBlue;
        }

        var headerCharacters = document.BeginUpdateCharacters(headerRow.Range);
        headerCharacters.Bold = true;
        headerCharacters.ForeColor = Color.White;
        document.EndUpdateCharacters(headerCharacters);

        var headerParagraphs = document.BeginUpdateParagraphs(headerRow.Range);
        headerParagraphs.Alignment = ParagraphAlignment.Center;
        document.EndUpdateParagraphs(headerParagraphs);

        for(int row = 1; row < table.Rows.Count; row++) {
            for(int col = 0; col < table.Rows[row].Cells.Count; col++) {
                var cell = table.Rows[row].Cells[col];
                var text = col == 0 ? $"Row {row + 1}" : $"Data {row + 1}-{col + 1}";
                document.InsertText(cell.Range.Start, text);
                cell.VerticalAlignment = TableCellVerticalAlignment.Center;
            }

            var cellParagraphs = document.BeginUpdateParagraphs(table.Rows[row].Range);
            cellParagraphs.Alignment = ParagraphAlignment.Center;
            document.EndUpdateParagraphs(cellParagraphs);
        }
    }
}
