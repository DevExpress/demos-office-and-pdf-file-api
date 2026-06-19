using DevExpress.Office.Services;
using DevExpress.Spreadsheet;
using DevExpress.Spreadsheet.Charts;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraSpreadsheet.Services;
using System.Drawing;
using RichEditDocumentFormat = DevExpress.XtraRichEdit.DocumentFormat;

class Program {
    static readonly string ChartTemplatePath = Path.Combine("Documents", "ChartTemplate.crtx");
    const string OutputFileName = "Chart_Document.docx";
    static readonly RichEditDocumentFormat OutputFormat = RichEditDocumentFormat.Docx;
    static readonly bool LoadFromTemplate = false;
    const bool ShowTrendline = true;
    const bool ShowDataLabels = true;
    const bool ShowAxisTitles = true;
    const bool ShowGridLines = true;
    const bool ShowLegend = true;
    const LegendPosition ChartLegendPosition = LegendPosition.Right;

    static int Main() {
        try {
            var outputPath = Path.Combine(AppContext.BaseDirectory, OutputFileName);
            var templatePath = Path.Combine(AppContext.BaseDirectory, ChartTemplatePath);
            using var chartTemplateStream = File.OpenRead(templatePath);
            using var outputStream = GenerateDocumentWithChart(chartTemplateStream, OutputFormat);
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

    static Stream GenerateDocumentWithChart(Stream chartTemplateInputStream, RichEditDocumentFormat outputFormat) {
        OfficeCharts.Instance.ActivateCrossPlatformCharts();

        using var wordProcessor = new RichEditDocumentServer();
        var document = wordProcessor.Document;
        document.Unit = DevExpress.Office.DocumentUnit.Inch;

        var chartShape = document.Shapes.InsertChart(document.Range.End, DevExpress.XtraRichEdit.API.Native.ChartType.ColumnClustered);
        chartShape.Size = new SizeF(6, 5);
        chartShape.RelativeHorizontalPosition = ShapeRelativeHorizontalPosition.Column;
        chartShape.RelativeVerticalPosition = ShapeRelativeVerticalPosition.Paragraph;
        chartShape.Offset = new PointF(0, 0);

        var chart = (ChartObject)chartShape.ChartFormat.Chart;
        var worksheet = (Worksheet)chartShape.ChartFormat.Worksheet;

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
        dataRange.Borders.SetOutsideBorders(Color.Black, DevExpress.Spreadsheet.BorderLineStyle.Thin);
        worksheet.Columns.AutoFit(0, 2);

        chart.SelectData(dataRange);
        chart.Series[1].ChangeType(DevExpress.Spreadsheet.Charts.ChartType.LineMarker);
        chart.Series[1].AxisGroup = AxisGroup.Secondary;
        chart.Title.SetValue("Example Chart");
        chart.Title.Visible = true;
        chart.Title.Font.Color = Color.FromArgb(0x1D, 0x2B, 0x64);
        chart.Legend.Visible = ShowLegend;
        chart.Legend.Position = ChartLegendPosition;
        chart.Views[0].DataLabels.ShowValue = ShowDataLabels;

        if(chart.PrimaryAxes.Count >= 2) {
            var xAxis = chart.PrimaryAxes[0];
            var yAxis = chart.PrimaryAxes[1];
            xAxis.Title.SetValue("Months");
            yAxis.Title.SetValue("Units");
            xAxis.Title.Visible = yAxis.Title.Visible = ShowAxisTitles;
            xAxis.MajorGridlines.Visible = yAxis.MajorGridlines.Visible = ShowGridLines;
        }

        if(ShowTrendline) {
            var series = chart.Series[0];
            var trendline = series.Trendlines.Add(ChartTrendlineType.Linear);
            trendline.CustomName = $"{series.SeriesName.PlainText} Trend";
            trendline.Outline.SetSolidFill(Color.Red);
        }

        if(LoadFromTemplate)
            chart.LoadTemplate(chartTemplateInputStream);

        var outputStream = new MemoryStream();
        wordProcessor.SaveDocument(outputStream, outputFormat);
        outputStream.Position = 0;
        return outputStream;
    }
}
