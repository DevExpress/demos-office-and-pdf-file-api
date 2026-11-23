using DevExpress.Docs.Spreadsheet;

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

    static Stream FilterByMultipleValues(Stream inputStream, IEnumerable<string> selectedProductFilterValues, DocumentFormat outputFormat) {
        using var workbook = new Workbook(inputStream);

        Worksheet regionalSalesWorksheet = GetRegionalSalesWorksheet(workbook);

        FilterValue[] filterValues;

        if(selectedProductFilterValues == null || !selectedProductFilterValues.Any()) {
            filterValues = ["Mozzarella di Giovanni", "Gorgonzola Telino"];
        }
        else {
            filterValues = selectedProductFilterValues.Select(value => (FilterValue)value).ToArray();
        }

        AutoFilter autoFilter = EnableFilter(regionalSalesWorksheet);

        AutoFilterColumn productsAutoFilterColumn = autoFilter.Columns[GetColumnIndexByName("Product")];
        productsAutoFilterColumn.Criteria = new ValueFilterCriteria(filterValues);

        return CreateOutputStream(workbook, outputFormat);
    }

    static Worksheet GetRegionalSalesWorksheet(Workbook workbook) {
        return workbook.Worksheets.First(sheet => sheet.Name == "Regional Sales");
    }

    static int GetColumnIndexByName(string columnName) {
        return columnName.ToLowerInvariant()
            switch {
                "region" => 0,
                "product" => 1,
                "sales" => 2,
                "reported date" => 3,
                _ => 0
            };
    }

    static AutoFilter EnableFilter(Worksheet worksheet) {
        worksheet.AutoFilter = new AutoFilter(worksheet["B2:E23"]);
        return worksheet.AutoFilter;
    }

    static Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
        var outputStream = new MemoryStream();

        workbook.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }
}
