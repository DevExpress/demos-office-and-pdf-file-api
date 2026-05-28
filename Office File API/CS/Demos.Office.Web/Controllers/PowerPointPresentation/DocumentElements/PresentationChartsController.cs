#region Code
using DevExpress.Docs.Presentation;
using DevExpress.Office.Utils;
using DevExpress.Docs.Office;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentElements;

namespace BlazorDemo.OfficeFileApi.Controllers.PowerPointPresentation.DocumentElements {
    [ApiController]
    [Route("api/presentation/charts")]
    public class PresentationChartsController : PresentationControllerBase {
        [HttpPost("generate")]
        public async Task<IActionResult> GenerateChart([FromForm] PresentationChartsModel model) {
            var outputStream = CreatePresentationChart(
                new ChartOptions() {
                    ChartTitle = model.ChartTitle,
                    ChartHeight = model.ChartHeight,
                    ChartStyle = model.ChartStyle,
                    ShowTrendline = model.ShowTrendline,
                    ShowDataLabels = model.ShowDataLabels,
                    ShowLegend = model.ShowLegend,
                    LegendPosition = model.LegendPosition,
                    ShowAxisTitles = model.ShowAxisTitles,
                    ShowGridLines = model.ShowGridLines
                },
                model.IsPdfOutput,
                PresentationFormatHelper.GetDocumentFormat(model.OutputFormat)
            );
            return CreateOutputFile(outputStream, model.IsPdfOutput ? "pdf" : model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream CreatePresentationChart(ChartOptions chartOptions, bool isPdfOutput, DocumentFormat format) {
            using var presentation = new Presentation();
            presentation.Slides.Clear();

            var slide = new Slide(SlideLayoutType.Blank);
            presentation.Slides.Add(slide);

            var chartHeightInDocuments = Units.InchesToDocumentsF(chartOptions.ChartHeight);
            var chart = new Chart(100, 100, chartHeightInDocuments * 2, chartHeightInDocuments);

            slide.Shapes.Add(chart);

            // Apply title settings
            var textArea = new TextArea(chartOptions.ChartTitle);
            textArea.Paragraphs.First().Properties.TextProperties.Fill = new SolidFill(Color.FromArgb(0x1D, 0x2B, 0x64));
            chart.Title = new ChartTitle(textArea);

            // Apply legend settings
            if(chartOptions.ShowLegend) {
                chart.Legend.Position = chartOptions.LegendPosition;
            } else {
                chart.Legend = null;
            }

            // Specify chart data
            var months = new[] { "January", "February", "March", "April", "May", "June" };
            var units = new[] { 50.0, 100.0, 30.0, 104.0, 87.0, 150.0 };
            var transactions = new[] { 900.0, 3000.0, 1200.0, 7000.0, 5100.0, 7500.0 };

            // Add primary series
            var unitsSoldSeries = new BarSeries {
                Text = "Units Sold",
                Arguments = new ChartStringData(months),
                Values = new ChartNumericData(units)
            };

            // Add trendline for the first series
            if(chartOptions.ShowTrendline) {
                unitsSoldSeries.TrendLine = new TrendLine() {
                    Name = $"{unitsSoldSeries.Text} trend",
                    TrendLineType = TrendLineType.Linear,
                    OutlineStyle = new LineStyle() {
                        Width = 2,
                        Fill = new SolidFill(Color.Red)
                    }
                };
            }

            // Show primary series data labels
            unitsSoldSeries.DataLabels = new BarDataLabels() {
                ShowValue = chartOptions.ShowDataLabels
            };
            chart.Series.Add(unitsSoldSeries);

            // Add secondary series
            var totalTransactionsSeries = new LineSeries {
                Text = "Total Transactions",
                Arguments = new ChartStringData(months),
                Values = new ChartNumericData(transactions),
                AxisGroup = ChartAxisGroupType.Secondary
            };
            chart.Series.Add(totalTransactionsSeries);

            // Apply primary axis settings
            if(chartOptions.ShowAxisTitles) {
                chart.ArgumentAxis.Title = new ChartTitleOptions("Months");
                chart.ValueAxis.Title = new ChartTitleOptions("Units");
            }

            // Apply secondary axis settings
            chart.SecondaryArgumentAxis.Visible = false;
            chart.SecondaryValueAxis.Position = AxisPositionType.Right;

            // Apply gridline settings
            ChartLineOptions gridLineOptions = new ChartLineOptions() {
                OutlineStyle = new OutlineStyle {
                    Fill = new SolidFill(Color.Black),
                }
            };
            chart.ValueAxis.MajorGridlines =
            chart.ArgumentAxis.MinorGridlines =
            chart.SecondaryValueAxis.MajorGridlines =
            chart.SecondaryArgumentAxis.MinorGridlines = chartOptions.ShowGridLines ? gridLineOptions : null;

            // Apply chart style
            chart.Style = chartOptions.ChartStyle;

            var outputStream = new MemoryStream();

            if(isPdfOutput)
                presentation.ExportToPdf(outputStream);
            else
                presentation.SaveDocument(outputStream, format);

            outputStream.Position = 0;
            return outputStream;
        }

        public class ChartOptions {
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
        #endregion Code
    }
}
