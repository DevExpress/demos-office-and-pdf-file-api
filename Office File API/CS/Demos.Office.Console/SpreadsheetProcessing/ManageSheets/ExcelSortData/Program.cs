using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "SortSample.xlsx");
    const string OutputFileName = "Sorted_Data.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = SortBySingleColumn(inputStream, "Continent", true, OutputFormat);
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

    static Stream SortBySingleColumn(Stream inputStream, string columnName, bool ascending, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets.First();
        CellRange range = worksheet.Range["A3:F22"];
        int columnIndex = GetColumnIndexByName(columnName);
        worksheet.Sort(range, columnIndex, ascending);
        worksheet["D1"].Value = $"Sort by column with index = {columnIndex} in {(ascending ? "ascending" : "descending")} order";
        return CreateOutputStream(workbook, outputFormat);
    }

    static int GetColumnIndexByName(string columnName) => columnName.ToLowerInvariant() switch {
        "continent" => 0,
        "country" => 1,
        "symbol" => 2,
        "population" => 3,
        "gdp" => 4,
        "area" => 5,
        _ => 0
    };

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
