using DevExpress.Docs;
using DevExpress.Docs.Word;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "ContentControls_AppointmentForm.docx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Table_Style.docx";
    const DocumentFormat OutputFormat = DocumentFormat.Docx;
    const TableLookTypes TableLook = TableLookTypes.ApplyFirstRow
                                   | TableLookTypes.DoNotApplyColumnBanding
                                   | TableLookTypes.DoNotApplyRowBanding;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = CreateApplyTableStyle(inputStream, TableLook, OutputFormat);
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

    static Stream CreateApplyTableStyle(Stream stream, TableLookTypes tableLookTypes, DocumentFormat outputFormat) {
        using var document = new WordDocument(stream);
        var table = document.Body.OfType<Table>().First();

        var tableStyle = CreateTableStyle("CustomTableStyle");
        // Add the style to the document collection
        document.TableStyles.Add(tableStyle);
        table.StyleName = tableStyle.Name;
        table.Look = tableLookTypes;

        var outputStream = new MemoryStream();
        document.Save(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static TableStyle CreateTableStyle(string styleName) {
        // Create a new table style
        var tableStyle = new TableStyle(styleName);

        // Specify style options
        var borders = tableStyle.TableProperties.Borders;
        var innerBorderColor = new OfficeColor(Color.White);
        borders.InsideHorizontal = new TableBorder(BorderLineStyle.Single, 0, innerBorderColor);
        borders.InsideVertical = new TableBorder(BorderLineStyle.Single, 0, innerBorderColor);

        tableStyle.TableCellProperties.BackColor = new OfficeColor(Color.FromArgb(227, 238, 220));

        var conditionalStyles = tableStyle.ConditionalStyles;
        conditionalStyles.Add(TableConditionalStyleType.FirstRow, CreateConditionalStyle(Color.FromArgb(173, 209, 158)));
        conditionalStyles.Add(TableConditionalStyleType.LastRow, CreateConditionalStyle(Color.FromArgb(195, 223, 179)));
        conditionalStyles.Add(TableConditionalStyleType.FirstColumn, CreateConditionalStyle(Color.FromArgb(210, 228, 200)));
        conditionalStyles.Add(TableConditionalStyleType.LastColumn, CreateConditionalStyle(Color.FromArgb(210, 228, 200)));
        conditionalStyles.Add(TableConditionalStyleType.OddRowBanding, CreateConditionalStyle(Color.FromArgb(235, 243, 229)));
        conditionalStyles.Add(TableConditionalStyleType.EvenRowBanding, CreateConditionalStyle(Color.FromArgb(255, 255, 255)));
        conditionalStyles.Add(TableConditionalStyleType.OddColumnBanding, CreateConditionalStyle(Color.FromArgb(239, 247, 233)));
        conditionalStyles.Add(TableConditionalStyleType.EvenColumnBanding, CreateConditionalStyle(Color.FromArgb(227, 238, 220)));
        return tableStyle;
    }

    static TableConditionalStyle CreateConditionalStyle(Color cellBackgroundColor) {
        var result = new TableConditionalStyle();
        result.TableCellProperties.BackColor = new OfficeColor(cellBackgroundColor);
        return result;
    }
}