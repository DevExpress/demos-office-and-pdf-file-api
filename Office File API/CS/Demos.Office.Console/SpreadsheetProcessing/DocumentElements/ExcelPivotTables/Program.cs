using DevExpress.Spreadsheet;
using System.Drawing;

class Program {
    const string OutputFileName = "PivotTable_Document.xlsx";
    static readonly PivotTableOptions Options = new() {
        CreatePivotTableSheet = false,
        ReportLayout = PivotReportLayout.Compact,
        ShowRowGrandTotals = true,
        ShowColumnGrandTotals = true,
        ShowSubtotals = true,
        ShowRowHeaders = true,
        ShowColumnHeaders = true,
        ShowFieldHeaders = true
    };
    const string SalesRegion = "East";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = GenerateDocumentWithPivotTable(Options, SalesRegion, OutputFormat);
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

    static Stream GenerateDocumentWithPivotTable(PivotTableOptions options, string salesRegion, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        Worksheet dataWorksheet = workbook.Worksheets.First();
        dataWorksheet.Name = "Pivot Data";

        dataWorksheet["A1"].Value = "Product";
        dataWorksheet["B1"].Value = "Region";
        dataWorksheet["C1"].Value = "Category";
        dataWorksheet["D1"].Value = "Sales";

        string[] products = ["Laptop", "Desktop", "Tablet", "Phone"];
        string[] regions = ["North", "South", "East", "West"];
        string[] categories = ["Electronics", "Gadgets", "Devices"];

        for(int i = 0; i < 20; i++) {
            int row = i + 2;
            dataWorksheet[$"A{row}"].Value = products[i % products.Length];
            dataWorksheet[$"B{row}"].Value = regions[i % regions.Length];
            dataWorksheet[$"C{row}"].Value = categories[i % categories.Length];
            dataWorksheet[$"D{row}"].Value = 1000 + (i * 150);
            dataWorksheet[$"D{row}"].NumberFormat = "$#,##0.00";
        }

        var headerRange = dataWorksheet["A1:D1"];
        headerRange.Font.Bold = true;
        headerRange.Fill.BackgroundColor = Color.LightGray;

        var dataRange = dataWorksheet.GetDataRange();
        dataRange.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);
        dataWorksheet.Columns.AutoFit(0, 3);

        Worksheet pivotWorksheet = dataWorksheet;
        if(options.CreatePivotTableSheet) {
            pivotWorksheet = workbook.Worksheets.Add("Pivot Table Sheet");
            workbook.Worksheets.ActiveWorksheet = pivotWorksheet;
        }

        PivotTable pivotTable = pivotWorksheet.PivotTables.Add(dataRange, options.CreatePivotTableSheet ? pivotWorksheet["A1"] : pivotWorksheet["F1"]);
        pivotTable.BeginUpdate();

        var regionField = pivotTable.RowFields.Add(pivotTable.Fields["Region"]);
        pivotTable.RowFields.Add(pivotTable.Fields["Product"]);
        pivotTable.ColumnFields.Add(pivotTable.Fields["Category"]);

        var salesField = pivotTable.DataFields.Add(pivotTable.Fields["Sales"]);
        salesField.NumberFormat = "$#,##0.00";
        pivotTable.Filters.Add(regionField.Field, PivotFilterType.CaptionEqual, salesRegion);

        pivotTable.Layout.SetReportLayout(options.ReportLayout);
        pivotTable.Layout.ShowRowGrandTotals = options.ShowRowGrandTotals;
        pivotTable.Layout.ShowColumnGrandTotals = options.ShowColumnGrandTotals;
        if(options.ShowSubtotals)
            pivotTable.Layout.ShowAllSubtotals(false);

        pivotTable.ShowRowHeaders = options.ShowRowHeaders;
        pivotTable.ShowColumnHeaders = options.ShowColumnHeaders;
        pivotTable.View.ShowFieldHeaders = options.ShowFieldHeaders;
        pivotTable.Style = workbook.TableStyles[BuiltInPivotStyleId.PivotStyleMedium2];
        pivotTable.EndUpdate();

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    sealed class PivotTableOptions {
        public bool CreatePivotTableSheet { get; set; }
        public PivotReportLayout ReportLayout { get; set; }
        public bool ShowRowGrandTotals { get; set; }
        public bool ShowColumnGrandTotals { get; set; }
        public bool ShowSubtotals { get; set; }
        public bool ShowRowHeaders { get; set; }
        public bool ShowColumnHeaders { get; set; }
        public bool ShowFieldHeaders { get; set; }
    }
}
