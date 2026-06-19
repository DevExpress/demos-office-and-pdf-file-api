using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using System.Drawing;

class Program {
    static readonly string ChartTemplatePath = Path.Combine("Documents", "ChartTemplate.crtx");
    const string OutputFileName = "Chart_Document.xlsx";
    static readonly ChartOptions Options = new() {
        LoadFromTemplate = false,
        CreateChartSheet = true,
        ShowTrendline = true,
        ShowDataLabels = true,
        ShowAxisTitles = true,
        ShowGridLines = true,
        ShowLegend = true,
        LegendPosition = LegendPosition.Right
    };
    static readonly DocumentFormat OutputFormat = DocumentFormat.Xlsx;

    static int Main() {
        try {
            var templatePath = Path.Combine(AppContext.BaseDirectory, ChartTemplatePath);
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);

            using var templateStream = File.OpenRead(templatePath);
            using var outputStream = GenerateDocumentWithChart(templateStream, Options, OutputFormat);
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

    static Stream GenerateDocumentWithChart(Stream chartTemplateInputStream, ChartOptions options, DocumentFormat outputFormat) {
        using var workbook = new Workbook();
        var worksheet = workbook.Worksheets.First();
        worksheet.Name = "Chart Data";

        var months = new[] { "January", "February", "March", "April", "May", "June" };
        var units = new[] { 50, 100, 30, 104, 87, 150 };
        var transactions = new[] { 900, 3000, 1200, 7000, 5100, 7500 };

        worksheet["A1"].Value = "Month";
        worksheet["B1"].Value = "Units Sold";
        worksheet["C1"].Value = "Total Transactions";

        for(int i = 0; i < months.Length; i++) {
            int row = i + 2;
            worksheet[$"A{row}"].Value = months[i];
            worksheet[$"B{row}"].Value = units[i];
            worksheet[$"C{row}"].Value = transactions[i];
        }

        var headerRange = worksheet["A1:C1"];
        headerRange.Font.Bold = true;
        headerRange.Fill.BackgroundColor = Color.LightGray;

        var dataRange = worksheet.GetDataRange();
        dataRange.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);
        worksheet.Columns.AutoFit(0, 2);

        ChartObject chart;
        if(options.CreateChartSheet) {
            var chartSheet = workbook.ChartSheets.Add(ChartType.ColumnClustered, dataRange);
            chartSheet.Name = "Chart Sheet";
            chart = chartSheet.Chart;
        }
        else {
            var chartInstance = worksheet.Charts.Add(ChartType.ColumnClustered, dataRange);
            chartInstance.TopLeftCell = worksheet.Cells["F1"];
            chartInstance.BottomRightCell = worksheet.Cells["R30"];
            chart = chartInstance;
        }

        chart.Series[1].ChangeType(ChartType.LineMarker);
        chart.Series[1].AxisGroup = AxisGroup.Secondary;
        chart.Title.SetValue("Example Chart");
        chart.Title.Visible = true;
        chart.Title.Font.Color = Color.FromArgb(0x1D, 0x2B, 0x64);
        chart.Legend.Visible = options.ShowLegend;
        chart.Legend.Position = options.LegendPosition;
        chart.Views[0].DataLabels.ShowValue = options.ShowDataLabels;

        if(chart.PrimaryAxes.Count >= 2) {
            var xAxis = chart.PrimaryAxes[0];
            var yAxis = chart.PrimaryAxes[1];
            xAxis.Title.SetValue("Months");
            yAxis.Title.SetValue("Units");
            xAxis.Title.Visible = yAxis.Title.Visible = options.ShowAxisTitles;
            xAxis.MajorGridlines.Visible = yAxis.MajorGridlines.Visible = options.ShowGridLines;
        }

        if(options.ShowTrendline) {
            var series = chart.Series[0];
            var trendline = series.Trendlines.Add(ChartTrendlineType.Linear);
            trendline.CustomName = $"{series.SeriesName.PlainText} Trend";
            trendline.Outline.SetSolidFill(Color.Red);
        }

        if(options.LoadFromTemplate)
            chart.LoadTemplate(chartTemplateInputStream);

        var outputStream = new MemoryStream();
        workbook.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }

    sealed class ChartOptions {
        public bool LoadFromTemplate { get; set; }
        public bool CreateChartSheet { get; set; }
        public bool ShowTrendline { get; set; }
        public bool ShowDataLabels { get; set; }
        public bool ShowAxisTitles { get; set; }
        public bool ShowGridLines { get; set; }
        public bool ShowLegend { get; set; }
        public LegendPosition LegendPosition { get; set; }
    }
}
