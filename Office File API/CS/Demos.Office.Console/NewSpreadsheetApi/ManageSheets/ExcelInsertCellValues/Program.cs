using DevExpress.Docs.Spreadsheet;
using DevExpress.Drawing;

class Program {
    static readonly string ImageFilePath = Path.Combine("Documents", "DevExpress.png");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Cell_Values.xlsx";
    const string TextValue = "Sample Text";
    const double NumericValue = 0.0;
    static readonly DateTime DateValue = DateTime.Today;
    const string NumberFormat = "0.00";
    const string DateFormat = "mm-dd-yy";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var imagePath = Path.Combine(AppContext.BaseDirectory, ImageFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var imageStream = File.OpenRead(imagePath);
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

    static Stream InsertCellValues(string textCellValue, double numericCellValue, DateTime dateCellValue,
        Stream imageStream, string numberFormat, string dateFormat, DocumentFormat outputFormat) {

        using var workbook = new Workbook();

        Worksheet firstWorksheet = workbook.Worksheets[0];
        firstWorksheet.Name = "Inserted Cell Values";

        firstWorksheet.Cells["A1"].Value = "Text Value";
        firstWorksheet.Cells["B1"].Value = "Numeric Value";
        firstWorksheet.Cells["C1"].Value = "Date Value";
        firstWorksheet.Cells["D1"].Value = "Image";

        firstWorksheet["A1:D1"].ApplyFont(new FontFormat(FontFormat.DefaultName, FontFormat.DefaultSize).WithBold(true));

        firstWorksheet.Cells["A2"].Value = textCellValue;

        firstWorksheet.Cells["B2"].Value = numericCellValue;
        firstWorksheet.Cells["B2"].Format.NumberFormat = numberFormat;

        firstWorksheet.Cells["C2"].SetDateValue(dateCellValue, dateFormat);

        firstWorksheet["A1:C2"].AutoFitColumns();

        firstWorksheet.Cells["D2"].SetImage(DXImage.FromStream(imageStream));

        firstWorksheet.Columns["D"].Width = 18;

        var outputStream = new MemoryStream();
        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
