using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "InvestmentPortfolio.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Modified_Document.xlsx";
    const string SearchText = "Technology";
    const string ReplaceText = "New Text";
    const FindReplaceActionType Operation = FindReplaceActionType.Replace;
    const bool MatchCase = false;
    const bool MatchEntireCellContent = false;
    static readonly Color HighlightColor = ColorTranslator.FromHtml("#FFFF00");
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = Operation switch {
                FindReplaceActionType.Replace => FindReplace(inputStream, SearchText, MatchCase, MatchEntireCellContent, ReplaceText, OutputFormat),
                FindReplaceActionType.Remove => FindRemove(inputStream, SearchText, MatchCase, MatchEntireCellContent, OutputFormat),
                FindReplaceActionType.Highlight => FindHighlight(inputStream, SearchText, MatchCase, MatchEntireCellContent, HighlightColor, OutputFormat),
                _ => FindReplace(inputStream, SearchText, MatchCase, MatchEntireCellContent, ReplaceText, OutputFormat)
            };
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

    static Stream FindReplace(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, string replaceText, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        var worksheet = workbook.Worksheets.ActiveWorksheet;
        var searchOptions = new SearchOptions() {
            SearchIn = SearchIn.Values,
            MatchCase = matchCase,
            MatchEntireCellContents = matchEntireCellContents
        };

        IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);
        foreach(Cell cell in searchResult)
            cell.SetValue(replaceText);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream FindRemove(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        var worksheet = workbook.Worksheets.ActiveWorksheet;
        var searchOptions = new SearchOptions() {
            SearchIn = SearchIn.Values,
            MatchCase = matchCase,
            MatchEntireCellContents = matchEntireCellContents
        };

        IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);
        foreach(Cell cell in searchResult)
            cell.Clear();

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream FindHighlight(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, Color highlightColor, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        workbook.LoadDocument(inputStream);

        var worksheet = workbook.Worksheets.ActiveWorksheet;
        var searchOptions = new SearchOptions() {
            SearchIn = SearchIn.Values,
            MatchCase = matchCase,
            MatchEntireCellContents = matchEntireCellContents
        };

        IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);
        foreach(Cell cell in searchResult)
            cell.Fill.BackgroundColor = highlightColor;

        return CreateOutputStream(workbook, outputFormat);
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}

enum FindReplaceActionType {
    Replace,
    Remove,
    Highlight
}
