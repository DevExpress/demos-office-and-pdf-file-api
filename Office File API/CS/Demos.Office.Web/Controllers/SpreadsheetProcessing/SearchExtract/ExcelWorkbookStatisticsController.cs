#region Code
using DevExpress.Spreadsheet;
using System.Text;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.SearchExtract {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelWorkbookStatisticsController : SpreadsheetControllerBase {
        [HttpPost("workbook-statistics/analyze")]
        public async Task<IActionResult> AnalyzeWorkbookStatistics(IFormFile file, [FromForm] ExcelWorkbookStatisticsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Seek(0, SeekOrigin.Begin);

            // Generate plain text statistics output
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
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Basic workbook statistics
            int worksheetCount = workbook.Worksheets.Count;
            int chartsheetCount = workbook.ChartSheets.Count;
            int visibleWorksheetCount = workbook.Worksheets.Count(ws => ws.Visible);

            output.AppendLine("BASIC WORKBOOK STATISTICS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Visible Worksheets: {visibleWorksheetCount}");
            output.AppendLine($"Total Worksheets: {worksheetCount}");
            output.AppendLine($"Total Chart Sheets: {chartsheetCount}");
            output.AppendLine();

            int totalDataCells = 0;
            int totalUsedCells = 0;

            int totalTables = 0;
            int totalPivotTables = 0;
            int totalCharts = 0;
            int totalImages = 0;
            int totalShapes = 0;
            int totalHyperlinks = 0;
            int totalComments = 0;

            // Worksheet statistics
            output.AppendLine("WORKSHEET STATISTICS:");
            output.AppendLine(new string('-', 35));

            foreach(var worksheet in workbook.Worksheets) {
                // Get statistics for visible worksheets
                if(skipHiddenWorksheets && !worksheet.Visible)
                    continue;

                var dataRange = worksheet.GetDataRange();
                var usedRange = worksheet.GetUsedRange();

                output.AppendLine(new string('-', 25));
                output.AppendLine($"Worksheet Index: {worksheet.Index}");
                output.AppendLine(new string('-', 25));

                output.AppendLine($"Name: {worksheet.Name}");
                output.AppendLine($"Visible: {(worksheet.Visible ? "Yes" : "No")}");
                output.AppendLine();

                output.AppendLine("CELL DATA:");
                int dataCellCount = dataRange.ExistingCells.Count();
                output.AppendLine($"Worksheet Data Range: {dataRange.GetReferenceA1()}");
                output.AppendLine($"Data Range Cell Count: {dataCellCount}");
                totalDataCells += dataCellCount;

                int usedCellCount = usedRange.ExistingCells.Count();
                output.AppendLine($"Worksheet Used Range: {usedRange.GetReferenceA1()}");
                output.AppendLine($"Used Range Cell Count: {usedCellCount}");
                totalUsedCells += usedCellCount;
                output.AppendLine();

                output.AppendLine("ELEMENTS:");
                int tablesCount = worksheet.Tables.Count;
                output.AppendLine($"Worksheet Tables Count: {tablesCount}");
                totalTables += tablesCount;

                int pivotTablesCount = worksheet.PivotTables.Count;
                output.AppendLine($"Worksheet Pivot Tables Count: {pivotTablesCount}");
                totalPivotTables += pivotTablesCount;

                int chartsCount = worksheet.Charts.Count;
                output.AppendLine($"Worksheet Charts Count: {chartsCount}");
                totalCharts += chartsCount;

                int imagesCount = worksheet.Pictures.Count;
                output.AppendLine($"Worksheet Images Count: {imagesCount}");
                totalImages += imagesCount;

                int shapesCount = worksheet.Shapes.Count;
                output.AppendLine($"Worksheet Shapes Count: {shapesCount}");
                totalShapes += shapesCount;

                int hyperlinksCount = worksheet.Hyperlinks.Count;
                output.AppendLine($"Worksheet Hyperlinks Count: {hyperlinksCount}");
                totalHyperlinks += hyperlinksCount;

                int commentsCount = worksheet.Comments.Count;
                output.AppendLine($"Worksheet Comments Count: {commentsCount}");
                totalComments += commentsCount;
                output.AppendLine();
            }

            // Workbook cells count
            output.AppendLine("WORKBOOK CELL DATA:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Total Data Cells: {totalDataCells}");
            output.AppendLine($"Total Used Cells: {totalUsedCells}");
            output.AppendLine();

            // Workbook elements count
            output.AppendLine("WORKBOOK ELEMENTS:");
            output.AppendLine(new string('-', 35));
            output.AppendLine($"Tables: {totalTables:N0}");
            output.AppendLine($"Pivot Tables: {totalPivotTables:N0}");
            output.AppendLine($"Charts: {totalCharts:N0}");
            output.AppendLine($"Images: {totalImages:N0}");
            output.AppendLine($"Shapes: {totalShapes:N0}");
            output.AppendLine($"Hyperlinks: {totalHyperlinks:N0}");
            output.AppendLine($"Comments: {totalComments:N0}");

            return output.ToString();
        }
        #endregion Code
    }
}
