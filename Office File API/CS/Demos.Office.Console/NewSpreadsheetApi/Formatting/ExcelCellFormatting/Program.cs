using System.Drawing;
using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "DocumentForProtection.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Cell_Formatting.xlsx";
    const bool CreateCustomStyle = false;
    const string FontName = "Segoe UI";
    const double FontSize = 16d;
    static readonly CellFontStyle FontStyle = CellFontStyle.Regular;
    static readonly Color FontColor = ColorTranslator.FromHtml("#000000");
    static readonly Color BackgroundColor = ColorTranslator.FromHtml("#66CDFF");
    const SpreadsheetHorizontalAlignment HorizontalAlignment = SpreadsheetHorizontalAlignment.General;
    const SpreadsheetVerticalAlignment VerticalAlignment = SpreadsheetVerticalAlignment.Center;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FormatWorkbook(inputStream, CreateCustomStyle, FontName, FontColor, FontSize,
                FontStyle is CellFontStyle.Bold or CellFontStyle.BoldItalic,
                FontStyle is CellFontStyle.Italic or CellFontStyle.BoldItalic,
                BackgroundColor, HorizontalAlignment, VerticalAlignment, OutputFormat);
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

    static Stream FormatWorkbook(Stream inputStream, bool createNewStyle, string fontName, Color fontColor, double fontSize,
        bool bold, bool italic, Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment,
        SpreadsheetVerticalAlignment verticalAlignment, DocumentFormat outputFormat) {

        using var workbook = new Workbook(inputStream);

        Worksheet worksheet = workbook.Worksheets[0];
        CellRange dataRange = worksheet.GetDataRange();
        if(dataRange.HasData) {

            var font = new FontFormat(fontName, fontSize, fontColor, bold, italic, false, UnderlineType.None);
            var fill = new SolidFillFormat(backgroundColor);
            var alignment = new AlignmentFormat(horizontalAlignment, verticalAlignment);

            if(createNewStyle) {
                Style customStyle = CreateStyle(workbook, font, fill, alignment);
                ApplyCustomStyle(dataRange, customStyle);
            }
            else {
                ApplyDirectFormatting(dataRange, font, fill, alignment);
            }

            dataRange.AutoFitColumns();
        }

        Stream outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static Style CreateStyle(Workbook workbook, FontFormat font, FillFormat fill, AlignmentFormat alignment) {
        var customStyle = new Style("Custom Style") {
            Font = font,
            Fill = fill,
            Alignment = alignment
        };

        workbook.Styles.Add(customStyle);
        return customStyle;
    }

    static void ApplyCustomStyle(CellRange cellRange, Style style) {
        cellRange.ApplyFont(null);
        cellRange.ApplyFill(null);
        cellRange.ApplyAlignment(null);

        cellRange.StyleName = style.Name;
    }

    static void ApplyDirectFormatting(CellRange cellRange, FontFormat font, FillFormat fill, AlignmentFormat alignment) {
        cellRange.ApplyFont(font);
        cellRange.ApplyFill(fill);
        cellRange.ApplyAlignment(alignment);
    }

    enum CellFontStyle {
        Regular,
        Italic,
        Bold,
        BoldItalic
    }
}
