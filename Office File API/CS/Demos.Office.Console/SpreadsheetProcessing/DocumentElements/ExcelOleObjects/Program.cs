using DevExpress.Spreadsheet;

class Program {
    static readonly string OleObjectPath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    static readonly string OleIconPath = Path.Combine("Documents", "Excel.png");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "OLE_Document.xlsx";
    const bool DisplayAsIcon = true;
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var oleObjectPath = Path.Combine(AppContext.BaseDirectory, OleObjectPath);
            var iconPath = Path.Combine(AppContext.BaseDirectory, OleIconPath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var oleObjectStream = File.OpenRead(oleObjectPath);
            using Stream? iconStream = DisplayAsIcon ? File.OpenRead(iconPath) : null;
            using var outputStream = GenerateWorkbookWithOleObject(oleObjectStream, OleObjectType.ExcelWorksheet, OutputFormat, iconStream);
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

    static Stream GenerateWorkbookWithOleObject(Stream oleObjectStream, OleObjectType oleObjectType, DocumentFormat outputFormat, Stream? oleIconStream = null) {
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Ole Object";

        string oleObjectRangeReference = "B2:N34";
        SpreadsheetImageSource? oleObjectIcon = null;

        if(oleIconStream != null) {
            oleObjectRangeReference = "B2:C6";
            oleObjectIcon = SpreadsheetImageSource.FromStream(oleIconStream);
        }

        var oleObjectRange = worksheet.Range[oleObjectRangeReference];
        worksheet.OleObjects.AddEmbeddedOleObject(oleObjectRange, oleObjectStream, oleObjectType, oleObjectIcon);

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
