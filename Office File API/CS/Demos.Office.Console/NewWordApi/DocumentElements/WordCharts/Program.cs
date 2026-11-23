using DevExpress.Docs.Office;
using DevExpress.Docs.Word;
using System.Drawing;

class Program {
    const string OutputDirectoryName = "Output";
    const string OutputFileName = "Chart_Document.docx";
    static readonly DocumentFormat OutputFormat = DocumentFormat.Docx;
    const bool ShowTrendline = true;
    const bool ShowDataLabels = true;
    const bool ShowAxisTitles = true;
    const bool ShowGridLines = true;
    const bool ShowLegend = true;
    const LegendPositionType LegendPosition = LegendPositionType.Right;

    static int Main() {
        try {
            var outputDirectory = Path.Combine(AppContext.BaseDirectory, OutputDirectoryName);
            Directory.CreateDirectory(outputDirectory);
            var outputPath = Path.Combine(outputDirectory, OutputFileName);
            var options = new ChartOptions {
                ShowTrendline = ShowTrendline,
                ShowDataLabels = ShowDataLabels,
                ShowAxisTitles = ShowAxisTitles,
                ShowGridLines = ShowGridLines,
                ShowLegend = ShowLegend,
                LegendPosition = LegendPosition
            };
            using var outputStream = GenerateDocumentWithChart(options, OutputFormat);
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

    static Stream GenerateDocumentWithChart(ChartOptions options, DocumentFormat outputFormat) {
        using var document = new WordDocument();

        // Create chart shape
        var chartShape = new ChartShape {
            // Specify the chart size
            Size = new SizeF(1800, 1500)
        };

        // Specify chart data
        var months = new[] { "January", "February", "March", "April", "May", "June" };
        var units = new[] { 50, 100, 30, 104, 87, 150 };
        var transactions = new[] { 900, 3000, 1200, 7000, 5100, 7500 };

        // Access a worksheet that stores chart data
        var chartData = chartShape.Data;
        chartData["A1"].TextValue = "Month";
        chartData["B1"].TextValue = "Units Sold";
        chartData["C1"].TextValue = "Total Transactions";

        for(var i = 0; i < months.Length; i++) {
            var row = i + 2;
            chartData[$"A{row}"].TextValue = months[i];
            chartData[$"B{row}"].NumericValue = units[i];
            chartData[$"C{row}"].NumericValue = transactions[i];
        }

        // Select chart data
        var unitsSoldSeries = new BarSeries {
            TextReference = new ChartDataExpressionReference("Sheet1!$B$1"),
            Arguments = new ChartDataExpressionReference("Sheet1!$A$2:$A$7"),
            Values = new ChartDataExpressionReference("Sheet1!$B$2:$B$7")
        };

        // Add trendline for the first series
        if(options.ShowTrendline) {
            unitsSoldSeries.TrendLine = new TrendLine {
                Name = "Units Sold Trend",
                Type = TrendLineType.Linear,
                OutlineStyle = new LineStyle { Fill = new SolidFill(Color.Red) }
            };
        }

        unitsSoldSeries.DataLabels.ShowValue = options.ShowDataLabels;
        chartShape.Series.Add(unitsSoldSeries);

        // Configure secondary series
        var totalTransactionsSeries = new LineSeries {
            TextReference = new ChartDataExpressionReference("Sheet1!$C$1"),
            Arguments = new ChartDataExpressionReference("Sheet1!$A$2:$A$7"),
            Values = new ChartDataExpressionReference("Sheet1!$C$2:$C$7"),
            Marker = new Marker { Symbol = MarkerStyle.Circle },
            AxisGroup = ChartAxisGroupType.Secondary
        };
        chartShape.Series.Add(totalTransactionsSeries);

        // Apply chart formatting
        var titleText = new TextArea("Example Chart");
        titleText.Paragraphs.First().Properties.TextProperties.Fill = new SolidFill(Color.FromArgb(0x1D, 0x2B, 0x64));
        chartShape.ChartTitle = new ChartTitle(titleText);

        if(options.ShowLegend) {
            chartShape.Legend.Position = options.LegendPosition;
        }
        else {
            chartShape.Legend = null;
        }

        // Apply primary axis settings
        if(options.ShowAxisTitles) {
            chartShape.ArgumentAxis.Title = new ChartTitleOptions("Months");
            chartShape.ValueAxis.Title = new ChartTitleOptions("Units");
        }

        chartShape.SecondaryArgumentAxis.Visible = false;
        chartShape.SecondaryValueAxis.Position = AxisPositionType.Right;

        chartShape.ArgumentAxis.MajorGridlines = options.ShowGridLines ? new ChartLineOptions() : null;
        chartShape.ValueAxis.MajorGridlines = options.ShowGridLines ? new ChartLineOptions() : null;

        var paragraph = document.Body.AddParagraph();
        paragraph.InlineElements.Add(chartShape);

        // Prepare output stream
        var outputStream = new MemoryStream();

        // Save the document
        document.Save(outputStream, outputFormat);

        outputStream.Position = 0;
        return outputStream;
    }

    public class ChartOptions {
        public bool ShowTrendline { get; set; }
        public bool ShowDataLabels { get; set; }
        public bool ShowAxisTitles { get; set; }
        public bool ShowGridLines { get; set; }
        public bool ShowLegend { get; set; }
        public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;
    }
}