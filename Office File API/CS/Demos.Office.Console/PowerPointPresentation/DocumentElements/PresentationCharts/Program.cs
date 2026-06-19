using System.Drawing;
using DevExpress.Docs.Office;
using DevExpress.Docs.Presentation;
using DevExpress.Office.Utils;

class Program {
    const string OutputFileName = "Presentation_Charts.pptx";
    const string OutputFormat = "pptx";
    const string ChartTitle = "Example Chart";
    const float ChartHeight = 5f;
    const bool ShowTrendline = true;
    const bool ShowDataLabels = true;
    const bool ShowAxisTitles = true;
    const bool ShowGridLines = true;
    const bool ShowLegend = true;
    const LegendPositionType LegendPosition = LegendPositionType.Right;
    const ChartStyle ChartStyleValue = ChartStyle.Color;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            using var outputStream = CreatePresentationChart(new ChartOptions {
                ChartTitle = ChartTitle,
                ChartHeight = ChartHeight,
                ChartStyle = ChartStyleValue,
                ShowTrendline = ShowTrendline,
                ShowDataLabels = ShowDataLabels,
                ShowLegend = ShowLegend,
                LegendPosition = LegendPosition,
                ShowAxisTitles = ShowAxisTitles,
                ShowGridLines = ShowGridLines
            }, IsPdfOutput(OutputFormat), GetDocumentFormat(OutputFormat));
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

    static Stream CreatePresentationChart(ChartOptions chartOptions, bool isPdfOutput, DocumentFormat format) {
        using var presentation = new Presentation();
        presentation.Slides.Clear();

        var slide = new Slide(SlideLayoutType.Blank);
        presentation.Slides.Add(slide);

        var chartHeightInDocuments = Units.InchesToDocumentsF(chartOptions.ChartHeight);
        var chart = new Chart(100, 100, chartHeightInDocuments * 2, chartHeightInDocuments);
        slide.Shapes.Add(chart);

        var textArea = new TextArea(chartOptions.ChartTitle);
        textArea.Paragraphs.First().Properties.TextProperties.Fill = new SolidFill(Color.FromArgb(0x1D, 0x2B, 0x64));
        chart.Title = new ChartTitle(textArea);

        if(chartOptions.ShowLegend)
            chart.Legend.Position = chartOptions.LegendPosition;
        else
            chart.Legend = null;

        var months = new[] { "January", "February", "March", "April", "May", "June" };
        var units = new[] { 50.0, 100.0, 30.0, 104.0, 87.0, 150.0 };
        var transactions = new[] { 900.0, 3000.0, 1200.0, 7000.0, 5100.0, 7500.0 };

        var unitsSoldSeries = new BarSeries {
            Text = "Units Sold",
            Arguments = new ChartStringData(months),
            Values = new ChartNumericData(units)
        };

        if(chartOptions.ShowTrendline) {
            unitsSoldSeries.TrendLine = new TrendLine {
                Name = $"{unitsSoldSeries.Text} trend",
                Type = TrendLineType.Linear,
                OutlineStyle = new LineStyle {
                    Width = 2,
                    Fill = new SolidFill(Color.Red)
                }
            };
        }

        unitsSoldSeries.DataLabels.ShowValue = chartOptions.ShowDataLabels;
        chart.Series.Add(unitsSoldSeries);

        var totalTransactionsSeries = new LineSeries {
            Text = "Total Transactions",
            Arguments = new ChartStringData(months),
            Values = new ChartNumericData(transactions),
            AxisGroup = ChartAxisGroupType.Secondary
        };
        chart.Series.Add(totalTransactionsSeries);

        if(chartOptions.ShowAxisTitles) {
            chart.ArgumentAxis.Title = new ChartTitleOptions("Months");
            chart.ValueAxis.Title = new ChartTitleOptions("Units");
        }

        chart.SecondaryArgumentAxis.Visible = false;
        chart.SecondaryValueAxis.Position = AxisPositionType.Right;

        ChartLineOptions gridLineOptions = new ChartLineOptions {
            OutlineStyle = new OutlineStyle {
                Fill = new SolidFill(Color.Black)
            }
        };
        chart.ValueAxis.MajorGridlines =
        chart.ArgumentAxis.MinorGridlines =
        chart.SecondaryValueAxis.MajorGridlines =
        chart.SecondaryArgumentAxis.MinorGridlines = chartOptions.ShowGridLines ? gridLineOptions : null;

        chart.Style = chartOptions.ChartStyle;

        var outputStream = new MemoryStream();
        if(isPdfOutput)
            presentation.ExportToPdf(outputStream);
        else
            presentation.SaveDocument(outputStream, format);

        outputStream.Position = 0;
        return outputStream;
    }

    static bool IsPdfOutput(string format) => format.Equals("pdf", StringComparison.OrdinalIgnoreCase);

    static DocumentFormat GetDocumentFormat(string format) {
        return format.ToLowerInvariant() switch {
            "pptm" => DocumentFormat.Pptm,
            "potx" => DocumentFormat.Potx,
            "potm" => DocumentFormat.Potm,
            _ => DocumentFormat.Pptx
        };
    }

    class ChartOptions {
        public string ChartTitle { get; set; } = "Example Chart";
        public float ChartHeight { get; set; } = 5f;
        public bool ShowTrendline { get; set; } = true;
        public bool ShowDataLabels { get; set; } = true;
        public bool ShowAxisTitles { get; set; } = true;
        public bool ShowGridLines { get; set; } = true;
        public bool ShowLegend { get; set; } = true;
        public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;
        public ChartStyle ChartStyle { get; set; } = ChartStyle.Color;
    }
}
