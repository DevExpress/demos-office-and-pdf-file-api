using DevExpress.Drawing.Printing;
using DevExpress.Office;
using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Print_Document.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = ApplyPrintSettings(inputStream, OutputFormat);
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

    static Stream ApplyPrintSettings(Stream inputStream, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        workbook.Unit = DocumentUnit.Inch;
        var worksheet = workbook.Worksheets.First();
        var activeView = worksheet.ActiveView;
        worksheet.SetPrintRange(worksheet.Range["A1:I30"]);
        activeView.Orientation = PageOrientation.Portrait;
        activeView.PaperKind = DXPaperKind.Letter;
        activeView.Margins.Top = 1.0f;
        activeView.Margins.Bottom = 1.0f;
        activeView.Margins.Left = 1.0f;
        activeView.Margins.Right = 1.0f;
        activeView.Margins.Header = 0.5f;
        activeView.Margins.Footer = 0.5f;
        var headerFooterOptions = worksheet.HeaderFooterOptions;
        headerFooterOptions.DifferentFirst = true;
        headerFooterOptions.FirstHeader.Left = string.Format("{0}&BDevExpress", HeaderFooterCode.FontColor(4, -50));
        headerFooterOptions.FirstHeader.Right = "&A";
        headerFooterOptions.FirstFooter.Left = "&D";
        headerFooterOptions.FirstFooter.Right = string.Format("Page {0} of {1}", "&P", "&N");
        headerFooterOptions.EvenHeader.Left = headerFooterOptions.OddHeader.Left = "Header Left";
        headerFooterOptions.EvenHeader.Center = headerFooterOptions.OddHeader.Center = "Header Center";
        headerFooterOptions.EvenHeader.Right = headerFooterOptions.OddHeader.Right = "Header Right";
        headerFooterOptions.EvenFooter.Left = headerFooterOptions.OddFooter.Left = "Footer Left";
        headerFooterOptions.EvenFooter.Center = headerFooterOptions.OddFooter.Center = "Footer Center";
        headerFooterOptions.EvenFooter.Right = headerFooterOptions.OddFooter.Right = "Footer Right";
        var printOptions = worksheet.PrintOptions;
        printOptions.FitToPage = true;
        printOptions.FitToWidth = 1;
        printOptions.FitToHeight = 1;
        printOptions.PrintGridlines = false;
        printOptions.PrintHeadings = false;
        printOptions.BlackAndWhite = false;
        printOptions.PageNumbering.Start = 1;
        printOptions.PageNumbering.StartType = WorksheetPageNumberingStartType.Restart;
        printOptions.PageOrder = PageOrder.DownThenOver;
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
