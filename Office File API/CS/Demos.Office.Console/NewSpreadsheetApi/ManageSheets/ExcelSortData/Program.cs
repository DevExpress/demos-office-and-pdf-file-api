using DevExpress.Docs.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "SortSample.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Sorted_Data.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);

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
        using var workbook = new Workbook(inputStream);

        Worksheet worksheet = workbook.Worksheets[0];

        CellRange range = worksheet["A3:F22"];
        int columnIndex = GetColumnIndexByName(columnName);
        worksheet.Sort(range, columnIndex, ascending);

        string order = ascending ? "ascending" : "descending";
        worksheet.Cells["D1"].Value = $"Sort by column with index = {columnIndex} in {order} order";

        return CreateOutputStream(workbook, outputFormat);
    }

    static int GetColumnIndexByName(string columnName) {
        return columnName.ToLowerInvariant()
            switch {
                "continent" => 0,
                "country" => 1,
                "symbol" => 2,
                "population" => 3,
                "gdp" => 4,
                "area" => 5,
                _ => 0
            };
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
