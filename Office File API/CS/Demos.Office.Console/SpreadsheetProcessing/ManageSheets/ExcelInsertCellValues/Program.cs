using DevExpress.Drawing;
using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    const string OutputFileName = "Cell_Values.xlsx";
    const string TextValue = "Sample Text";
    const double NumericValue = 0.0;
    static readonly DateTime DateValue = DateTime.Today;
    static readonly string NumberFormat = "0.00";
    static readonly string DateFormat = "m/d/yyyy";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var imageStream = CreateSampleImageStream();
            using var outputStream = InsertCellValues(TextValue, NumericValue, DateValue, imageStream, NumberFormat, DateFormat, OutputFormat);
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

    static Stream CreateSampleImageStream() {
        var stream = new MemoryStream();
        using var bitmap = new DXBitmap(192, 32);
        using(DXGraphics graphics = DXGraphics.FromImage(bitmap)) {
            graphics.Clear(Color.White);
            using var brush = new DXSolidBrush(Color.DeepSkyBlue);
            graphics.FillRectangle(brush, 0, 0, 192, 32);
            graphics.DrawString(
                "DevExpress",
                new DXFont("Arial", 14, DXFontStyle.Bold),
                DXBrushes.White,
                new PointF(28, 6),
                DXStringFormat.CreateGenericTypographic());
        }
        bitmap.Save(stream, DXImageFormat.Png);
        stream.Position = 0;
        return stream;
    }

    static Stream InsertCellValues(string textCellValue, double numericCellValue, DateTime dateCellValue, Stream imageStream, string numberFormat, string dateFormat, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.Unit = DevExpress.Office.DocumentUnit.Inch;
        Worksheet worksheet = workbook.Worksheets.First();
        worksheet.Name = "Inserted Cell Values";
        worksheet["A1"].Value = "Text Value";
        worksheet["B1"].Value = "Numeric Value";
        worksheet["C1"].Value = "Date Value";
        worksheet["D1"].Value = "Image";
        var formatting = worksheet["A1:D1"].BeginUpdateFormatting();
        formatting.Font.Bold = true;
        worksheet["A1:D1"].EndUpdateFormatting(formatting);
        worksheet["A2"].Value = textCellValue;
        worksheet["B2"].Value = numericCellValue;
        worksheet["B2"].NumberFormat = numberFormat;
        worksheet["C2"].Value = dateCellValue;
        worksheet["C2"].NumberFormat = dateFormat;
        worksheet["D2"].Value = DXImage.FromStream(imageStream);
        worksheet.Columns.AutoFit(0, 4);
        worksheet.Columns["D"].Width = 1;
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
