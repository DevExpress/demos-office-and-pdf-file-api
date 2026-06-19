using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputFileName = "Calculated_Document.pdf";
    static readonly string OutputFormat = "pdf";
    static readonly CalculationMode? DocumentCalculationMode = null;
    static readonly bool ExportToPdf = true;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = CalculateWorkbook(inputStream, DocumentCalculationMode, ExportToPdf, GetDocumentFormat(OutputFormat));
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

    static Stream CalculateWorkbook(Stream inputStream, CalculationMode? documentCalculationMode, bool exportToPdf, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();

        workbook.LoadDocument(inputStream);
        workbook.Calculate();

        if(documentCalculationMode.HasValue)
            workbook.DocumentSettings.Calculation.Mode = documentCalculationMode.Value;

        var outputStream = new MemoryStream();
        if(exportToPdf)
            workbook.ExportToPdf(outputStream);
        else
            workbook.SaveDocument(outputStream, outputFormat);

        outputStream.Seek(0, SeekOrigin.Begin);
        return outputStream;
    }

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "xls" => DocumentFormat.Xls,
            "xlsx" => DocumentFormat.Xlsx,
            "xlt" => DocumentFormat.Xlt,
            "xlsm" => DocumentFormat.Xlsm,
            "xltx" => DocumentFormat.Xltx,
            "xltm" => DocumentFormat.Xltm,
            "xlsb" => DocumentFormat.Xlsb,
            "xml" => DocumentFormat.XmlSpreadsheet2003,
            "csv" => DocumentFormat.Csv,
            "txt" => DocumentFormat.Text,
            _ => DocumentFormat.Xlsx
        };
    }
}
