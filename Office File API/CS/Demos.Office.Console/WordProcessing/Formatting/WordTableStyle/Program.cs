using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "ContentControls_AppointmentForm.docx");
    const string OutputFileName = "Table_Style.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;

    const TableLookTypes TableLook = TableLookTypes.ApplyFirstRow
                                   | TableLookTypes.ApplyFirstColumn
                                   | TableLookTypes.DoNotApplyRowBanding;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
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
        using var wordProcessor = new RichEditDocumentServer();
        wordProcessor.LoadDocument(stream);
        var document = wordProcessor.Document;
        var table = document.Tables[0];
        table.Style = CreateTableStyle(document, "CustomTableStyle");
        table.TableLook = tableLookTypes;

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    static TableStyle CreateTableStyle(Document document, string styleName) {
        var tableStyle = document.TableStyles.CreateNew();
        tableStyle.Name = styleName;

        var tableBorders = tableStyle.TableBorders;
        tableBorders.InsideHorizontalBorder.LineStyle = BorderLineStyle.Single;
        tableBorders.InsideHorizontalBorder.LineColor = Color.White;
        tableBorders.InsideVerticalBorder.LineStyle = BorderLineStyle.Single;
        tableBorders.InsideVerticalBorder.LineColor = Color.White;
        tableStyle.CellBackgroundColor = Color.FromArgb(227, 238, 220);

        var conditionalStyleProperties = tableStyle.ConditionalStyleProperties;
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstRow).CellBackgroundColor = Color.FromArgb(173, 209, 158);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.LastRow).CellBackgroundColor = Color.FromArgb(195, 223, 179);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.FirstColumn).CellBackgroundColor = Color.FromArgb(210, 228, 200);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.LastColumn).CellBackgroundColor = Color.FromArgb(210, 228, 200);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddRowBanding).CellBackgroundColor = Color.FromArgb(235, 243, 229);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenRowBanding).CellBackgroundColor = Color.White;
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.OddColumnBanding).CellBackgroundColor = Color.FromArgb(239, 247, 233);
        conditionalStyleProperties.CreateConditionalStyle(ConditionalTableStyleFormattingTypes.EvenColumnBanding).CellBackgroundColor = Color.FromArgb(227, 238, 220);

        document.TableStyles.Add(tableStyle);
        return tableStyle;
    }
}
