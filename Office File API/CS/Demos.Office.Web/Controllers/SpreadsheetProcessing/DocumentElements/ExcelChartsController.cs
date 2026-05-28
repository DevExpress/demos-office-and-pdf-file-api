#region Code
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using DevExpress.XtraSpreadsheet.Services;
using DevExpress.Office.Services;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelChartsController : SpreadsheetControllerBase {
        [HttpPost("charts/generate")]
        public async Task<IActionResult> GenerateChartWorkbook([FromForm] IFormFile file, [FromForm] ExcelChartsModel model) {
            // Read file stream
            using var chartTemplateInputStream = new MemoryStream();
            await file.CopyToAsync(chartTemplateInputStream);
            chartTemplateInputStream.Position = 0;

            // Determine target format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with a chart
            var outputStream = GenerateDocumentWithChart(
                chartTemplateInputStream,
                new ChartOptions {
                    LoadFromTemplate = model.LoadFromTemplate,
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
        Stream GenerateDocumentWithChart(Stream chartTemplateInputStream, ChartOptions options, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            var worksheet = workbook.Worksheets.First();
            worksheet.Name = "Chart Data";

            // Specify chart data
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

            // Apply formatting to chart data cells
            var headerRange = worksheet["A1:C1"];
            headerRange.Font.Bold = true;
            headerRange.Fill.BackgroundColor = Color.LightGray;

            var dataRange = worksheet.GetDataRange();
            dataRange.Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thin);

            worksheet.Columns.AutoFit(0, 2);

            // Create chart
            ChartObject chart;

            if(options.CreateChartSheet) {
                var chartSheet = workbook.ChartSheets.Add(ChartType.ColumnClustered, dataRange);
                chartSheet.Name = "Chart Sheet";
                chart = chartSheet.Chart;
            } else {
                var chartInstance = worksheet.Charts.Add(ChartType.ColumnClustered, dataRange);
                // For an embedded chart, set the location and size
                chartInstance.TopLeftCell = worksheet.Cells["F1"];
                chartInstance.BottomRightCell = worksheet.Cells["R30"];
                chart = chartInstance;
            }

            // Configure secondary series
            chart.Series[1].ChangeType(DevExpress.Spreadsheet.Charts.ChartType.LineMarker);
            chart.Series[1].AxisGroup = AxisGroup.Secondary;

            // Apply chart formatting
            chart.Title.SetValue("Example Chart");
            chart.Title.Visible = true;
            chart.Title.Font.Color = Color.FromArgb(0x1D, 0x2B, 0x64);

            chart.Legend.Visible = options.ShowLegend;
            chart.Legend.Position = options.LegendPosition;

            chart.Views[0].DataLabels.ShowValue = options.ShowDataLabels;

            // Apply primary axis settings
            if(chart.PrimaryAxes.Count >= 2) {
                var xAxis = chart.PrimaryAxes[0];
                var yAxis = chart.PrimaryAxes[1];

                xAxis.Title.SetValue("Months");
                yAxis.Title.SetValue("Units");

                xAxis.Title.Visible = yAxis.Title.Visible = options.ShowAxisTitles;
                xAxis.MajorGridlines.Visible = yAxis.MajorGridlines.Visible = options.ShowGridLines;
            }

            // Add trendline for the first series
            if(options.ShowTrendline) {
                var series = chart.Series[0];
                var trendline = series.Trendlines.Add(ChartTrendlineType.Linear);

                trendline.CustomName = $"{series.SeriesName.PlainText} Trend";
                trendline.Outline.SetSolidFill(Color.Red);
            }

            // Load chart template
            if(options.LoadFromTemplate) {
                chart.LoadTemplate(chartTemplateInputStream);
            }

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }

        public class ChartOptions {
            public bool LoadFromTemplate { get; set; }
            public bool CreateChartSheet { get; set; }
            public bool ShowTrendline { get; set; }
            public bool ShowDataLabels { get; set; }
            public bool ShowAxisTitles { get; set; }
            public bool ShowGridLines { get; set; }
            public bool ShowLegend { get; set; }
            public LegendPosition LegendPosition { get; set; } = LegendPosition.Right;
        }
        #endregion Code
    }
}
