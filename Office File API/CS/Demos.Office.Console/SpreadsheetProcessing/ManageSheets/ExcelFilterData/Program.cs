using DevExpress.Spreadsheet;

class Program {
    static readonly string InputFilePath = Path.Combine("Documents", "FilterSample.xlsx");
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Filtered_Data.xlsx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var inputPath = Path.Combine(AppContext.BaseDirectory, InputFilePath);
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            using var inputStream = File.OpenRead(inputPath);
            using var outputStream = FilterByMultipleValues(inputStream, ["Mozzarella di Giovanni", "Gorgonzola Telino"], OutputFormat);
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

    static Stream FilterByMultipleValues(Stream inputStream, IEnumerable<string> values, DocumentFormat outputFormat) {
        using Workbook workbook = new Workbook();
        workbook.LoadDocument(inputStream);
        Worksheet worksheet = workbook.Worksheets["Regional sales"];
        EnableFilter(worksheet);
        CellValue[] cellValues = values.Select(v => (CellValue)v).ToArray();
        worksheet.AutoFilter.Columns[GetColumnIndexByName("Product")].ApplyFilterCriteria(cellValues);
        return CreateOutputStream(workbook, outputFormat);
    }

    static void EnableFilter(Worksheet worksheet) {
        worksheet.AutoFilter.Apply(worksheet["B2:E23"]);
    }

    static int GetColumnIndexByName(string columnName) => columnName.ToLowerInvariant() switch {
        "region" => 0,
        "product" => 1,
        "sales" => 2,
        "reported date" => 3,
        _ => 0
    };

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
