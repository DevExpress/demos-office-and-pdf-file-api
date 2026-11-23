#region Code
using DevExpress.Docs.Spreadsheet;
using DevExpress.Docs.Office;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.DocumentElements;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.DocumentElements {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelChartsController : SpreadsheetControllerBase {
        [HttpPost("charts/generate")]
        public async Task<IActionResult> GenerateChartWorkbook([FromForm] ExcelChartsModel model) {
            var outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var outputStream = GenerateDocumentWithChart(
                new ChartOptions {
                    CreateChartSheet = model.CreateChartSheet,
                    ShowTrendline = model.ShowTrendline,
                    ShowDataLabels = model.ShowDataLabels,
                    ShowAxisTitles = model.ShowAxisTitles,
                    ShowGridLines = model.ShowGridLines,
                    ShowLegend = model.ShowLegend,
                    LegendPosition = model.LegendPosition,
                }, outputFormat);

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GenerateDocumentWithChart(ChartOptions options, DocumentFormat outputFormat) {
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
            headerRange.ApplyFont(new FontFormat().WithBold(true));
            headerRange.ApplyFill(new SolidFillFormat(Color.LightGray));

            var dataRange = worksheet["A1:C7"];
            dataRange.ApplyBorders(new BorderFormat(new BorderLine(BorderLineStyle.Thin, Color.Black)));
            dataRange.AutoFitColumns();

            var chart = new Chart();

            var unitsSoldSeries = new BarSeries {
                TextReference = worksheet["B1"].GetExpressionReference(),
                Arguments = worksheet["A2:A7"].GetExpressionReference(),
                Values = worksheet["B2:B7"].GetExpressionReference()
            };

            if(options.ShowTrendline) {
                unitsSoldSeries.TrendLine = new TrendLine {
                    Name = "Units Sold Trend",
                    Type = TrendLineType.Linear,
                    OutlineStyle = new LineStyle { Fill = new SolidFill(Color.Red) }
                };
            }

            unitsSoldSeries.DataLabels.ShowValue = options.ShowDataLabels;
            chart.Series.Add(unitsSoldSeries);

            var totalTransactionsSeries = new LineSeries {
                TextReference = worksheet["C1"].GetExpressionReference(),
                Arguments = worksheet["A2:A7"].GetExpressionReference(),
                Values = worksheet["C2:C7"].GetExpressionReference(),
                Marker = new Marker { Symbol = MarkerStyle.Circle },
                AxisGroup = ChartAxisGroupType.Secondary
            };
            chart.Series.Add(totalTransactionsSeries);

            if(options.CreateChartSheet) {
                workbook.ChartSheets.Add(new ChartSheet("Chart Sheet", chart));
            }
            else {
                chart.ShapeAnchor = new TwoCellAnchor(0, 5, 30, 18);
                worksheet.Shapes.Add(chart);
            }

            var titleText = new TextArea("Example Chart");
            titleText.Paragraphs.First().Properties.TextProperties.Fill = new SolidFill(Color.FromArgb(0x1D, 0x2B, 0x64));
            chart.Title = new ChartTitle(titleText);

            if(options.ShowLegend)
                chart.Legend.Position = options.LegendPosition;
            else
                chart.Legend = null;

            if(options.ShowAxisTitles) {
                chart.ArgumentAxis.Title = new ChartTitleOptions("Months");
                chart.ValueAxis.Title = new ChartTitleOptions("Units");
            }

            chart.SecondaryArgumentAxis.Visible = false;
            chart.SecondaryValueAxis.Position = AxisPositionType.Right;

            chart.ArgumentAxis.MajorGridlines = options.ShowGridLines ? new ChartLineOptions() : null;
            chart.ValueAxis.MajorGridlines = options.ShowGridLines ? new ChartLineOptions() : null;

            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        public class ChartOptions {
            public bool CreateChartSheet { get; set; }
            public bool ShowTrendline { get; set; }
            public bool ShowDataLabels { get; set; }
            public bool ShowAxisTitles { get; set; }
            public bool ShowGridLines { get; set; }
            public bool ShowLegend { get; set; }
            public LegendPositionType LegendPosition { get; set; } = LegendPositionType.Right;
        }
        #endregion Code
    }
}
