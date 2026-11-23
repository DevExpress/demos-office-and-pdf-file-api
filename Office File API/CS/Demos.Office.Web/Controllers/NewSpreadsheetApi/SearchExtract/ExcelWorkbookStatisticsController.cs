#region Code
using DevExpress.Docs.Spreadsheet;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.SearchExtract;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.SearchExtract {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelWorkbookStatisticsController : SpreadsheetControllerBase {
        [HttpPost("workbook-statistics/analyze")]
        public async Task<IActionResult> AnalyzeWorkbookStatistics(IFormFile file, [FromForm] ExcelWorkbookStatisticsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            var statisticsText = GenerateWorkbookStatisticsText(inputStream, model.SkipHiddenWorksheets);

            using var outputStream = new MemoryStream();
            using(StreamWriter writer = new StreamWriter(outputStream, Encoding.UTF8)) {
                writer.Write(statisticsText);
                writer.Flush();
                outputStream.Position = 0;
                return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
            }
        }

        #region Code
        string GenerateWorkbookStatisticsText(Stream inputStream, bool skipHiddenWorksheets) {
            var output = new StringBuilder();
            using var workbook = new Workbook(inputStream);

            int worksheetCount = workbook.Worksheets.Count;
            int chartsheetCount = workbook.ChartSheets.Count;
            int visibleWorksheetCount = workbook.Worksheets.Count(ws => ws.Visibility == SheetVisibility.Visible);

            output.AppendLine("BASIC WORKBOOK STATISTICS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Visible Worksheets: {visibleWorksheetCount}");
            output.AppendLine($"Total Worksheets: {worksheetCount}");
            output.AppendLine($"Total Chart Sheets: {chartsheetCount}");
            output.AppendLine();

            int totalUsedCells = 0;

            int totalTables = 0;
            int totalCharts = 0;
            int totalImages = 0;
            int totalShapes = 0;

            output.AppendLine("WORKSHEET STATISTICS:");
            output.AppendLine(new string('-', 35));

            for(int index = 0; index < workbook.Worksheets.Count; index++) {
                Worksheet worksheet = workbook.Worksheets[index];
                bool isVisible = worksheet.Visibility == SheetVisibility.Visible;

                if(skipHiddenWorksheets && !isVisible)
                    continue;

                output.AppendLine(new string('-', 25));
                output.AppendLine($"Worksheet Index: {index}");
                output.AppendLine(new string('-', 25));

                output.AppendLine($"Name: {worksheet.Name}");
                output.AppendLine($"Visible: {(isVisible ? "Yes" : "No")}");
                output.AppendLine();

                output.AppendLine("CELL DATA:");
                int usedCellCount = worksheet.Cells.Count;
                output.AppendLine($"Used Cell Count: {usedCellCount}");
                output.AppendLine($"Used Range: {worksheet.GetUsedRange()}");
                output.AppendLine($"Data Range: {worksheet.GetDataRange()}");
                totalUsedCells += usedCellCount;
                output.AppendLine();

                output.AppendLine("ELEMENTS:");
                int tablesCount = worksheet.Tables.Count;
                output.AppendLine($"Worksheet Tables Count: {tablesCount}");
                totalTables += tablesCount;

                int chartsCount = worksheet.Shapes.OfType<ChartBase>().Count();
                output.AppendLine($"Worksheet Charts Count: {chartsCount}");
                totalCharts += chartsCount;

                int imagesCount = worksheet.Shapes.GetImageShapes().Count();
                output.AppendLine($"Worksheet Images Count: {imagesCount}");
                totalImages += imagesCount;

                int shapesCount = worksheet.Shapes.Count;
                output.AppendLine($"Worksheet Shapes Count: {shapesCount}");
                totalShapes += shapesCount;
                output.AppendLine();
            }

            output.AppendLine("WORKBOOK CELL DATA:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Total Used Cells: {totalUsedCells}");
            output.AppendLine();

            output.AppendLine("WORKBOOK ELEMENTS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Tables: {totalTables:N0}");
            output.AppendLine($"Charts: {totalCharts:N0}");
            output.AppendLine($"Images: {totalImages:N0}");
            output.AppendLine($"Shapes: {totalShapes:N0}");

            return output.ToString();
        }
        #endregion Code
    }
}
