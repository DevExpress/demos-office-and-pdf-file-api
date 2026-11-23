using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet.Export;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "ConvertedSheet.html";
    const int SheetIndex = 0;
    const bool ExportImages = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ConvertToHtml(inputStream, SheetIndex, ExportImages);
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

    static Stream ConvertToHtml(Stream inputStream, int sheetIndex, bool exportImages) {
        using var workbook = new Workbook();

        workbook.LoadDocument(inputStream);

        var outputStream = new MemoryStream();
        var exportOptions = new HtmlDocumentExporterOptions {
            SheetIndex = sheetIndex,
            EmbedImages = true,
            ExportImages = exportImages
        };

        workbook.ExportToHtml(outputStream, exportOptions);
        outputStream.Position = 0;
        return outputStream;
    }
}
