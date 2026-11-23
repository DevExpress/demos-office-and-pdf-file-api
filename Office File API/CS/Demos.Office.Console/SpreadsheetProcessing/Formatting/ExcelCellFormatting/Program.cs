using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "DocumentForProtection.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Cell_Formatting.xlsx";
    const bool CreateCustomStyle = false;
    const string FontName = "Segoe UI";
    const double FontSize = 16d;
    const SpreadsheetFontStyle FontStyle = SpreadsheetFontStyle.Regular;
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
            using var outputStream = FormatWorkbook(inputStream, CreateCustomStyle, FontName, FontColor, FontSize, FontStyle, BackgroundColor, HorizontalAlignment, VerticalAlignment, OutputFormat);
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
        SpreadsheetFontStyle fontStyle, Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment,
        SpreadsheetVerticalAlignment verticalAlignment, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        CellRange dataRange = workbook.Worksheets.ActiveWorksheet.GetDataRange();

        if(createNewStyle) {
            Style customStyle = CreateStyle(workbook, fontName, fontColor, fontSize, fontStyle, backgroundColor, horizontalAlignment, verticalAlignment);
            dataRange.Style = customStyle;
        }
        else {
            ApplyDirectFormatting(dataRange, fontName, fontColor, fontSize, fontStyle, backgroundColor, horizontalAlignment, verticalAlignment);
        }

        dataRange.AutoFitColumns();

        Stream outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static Style CreateStyle(Workbook workbook, string fontName, Color fontColor, double fontSize, SpreadsheetFontStyle fontStyle,
        Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment, SpreadsheetVerticalAlignment verticalAlignment) {
        Style customStyle = workbook.Styles.Add("Custom Style");
        customStyle.BeginUpdate();
        try {
            customStyle.Font.Name = fontName;
            customStyle.Font.Color = fontColor;
            customStyle.Font.Size = fontSize;
            customStyle.Font.FontStyle = fontStyle;
            customStyle.Fill.BackgroundColor = backgroundColor;
            customStyle.Alignment.Horizontal = horizontalAlignment;
            customStyle.Alignment.Vertical = verticalAlignment;
        }
        finally {
            customStyle.EndUpdate();
        }

        return customStyle;
    }

    static void ApplyDirectFormatting(CellRange cellRange, string fontName, Color fontColor, double fontSize, SpreadsheetFontStyle fontStyle,
        Color backgroundColor, SpreadsheetHorizontalAlignment horizontalAlignment, SpreadsheetVerticalAlignment verticalAlignment) {
        var formatting = cellRange.BeginUpdateFormatting();
        try {
            formatting.Font.Name = fontName;
            formatting.Font.Color = fontColor;
            formatting.Font.Size = fontSize;
            formatting.Font.FontStyle = fontStyle;
            formatting.Fill.BackgroundColor = backgroundColor;
            formatting.Alignment.Horizontal = horizontalAlignment;
            formatting.Alignment.Vertical = verticalAlignment;
        }
        finally {
            cellRange.EndUpdateFormatting(formatting);
        }
    }
}
