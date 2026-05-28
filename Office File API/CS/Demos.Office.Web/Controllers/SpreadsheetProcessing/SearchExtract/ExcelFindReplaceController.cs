#region Code
using DevExpress.Spreadsheet;
using System.Drawing;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.SearchExtract {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelFindReplaceController : SpreadsheetControllerBase {
        [HttpPost("find-replace/process")]
        public async Task<IActionResult> ProcessFindReplace(IFormFile file, [FromForm] ExcelFindReplaceModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            Stream outputStream = new MemoryStream();
            switch(model.Operation) {
                case FindReplaceActionType.Replace:
                    outputStream = FindReplace(inputStream, model.SearchText, model.MatchCase, model.MatchEntireCellContent, model.ReplaceText ?? "", targetFormat);
                    break;
                case FindReplaceActionType.Remove:
                    outputStream = FindRemove(inputStream, model.SearchText, model.MatchCase, model.MatchEntireCellContent, targetFormat);
                    break;
                case FindReplaceActionType.Highlight:
                    outputStream = FindHighlight(inputStream, model.SearchText, model.MatchCase, model.MatchEntireCellContent, ColorTranslator.FromHtml(model.HighlightColor), targetFormat);
                    break;
            }
            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream FindReplace(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, string replaceText, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            // Load document
            workbook.LoadDocument(inputStream);

            // Search text in an active worksheet
            var worksheet = workbook.Worksheets.ActiveWorksheet;

            // Specify search options
            var searchOptions = new SearchOptions() {
                SearchIn = SearchIn.Values,
                MatchCase = matchCase,
                MatchEntireCellContents = matchEntireCellContents
            };

            // Find all cells that contain the specified search text
            IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);

            // Replace text in found cells
            foreach(Cell cell in searchResult)
                cell.SetValue(replaceText);

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FindRemove(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            // Load document
            workbook.LoadDocument(inputStream);

            // Search text in an active worksheet
            var worksheet = workbook.Worksheets.ActiveWorksheet;

            // Specify search options
            var searchOptions = new SearchOptions() {
                SearchIn = SearchIn.Values,
                MatchCase = matchCase,
                MatchEntireCellContents = matchEntireCellContents
            };

            // Find all cells that contain the specified search text
            IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);

            // Clear content in found cells
            foreach(Cell cell in searchResult)
                cell.Clear();

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream FindHighlight(Stream inputStream, string searchText, bool matchCase, bool matchEntireCellContents, Color highlightColor, DocumentFormat outputFormat) {
            using var workbook = new Workbook();

            // Load document
            workbook.LoadDocument(inputStream);

            // Search text in an active worksheet
            var worksheet = workbook.Worksheets.ActiveWorksheet;

            // Specify search options
            var searchOptions = new SearchOptions() {
                SearchIn = SearchIn.Values,
                MatchCase = matchCase,
                MatchEntireCellContents = matchEntireCellContents
            };

            // Find all cells that contain the specified search text
            IEnumerable<Cell> searchResult = worksheet.Search(searchText, searchOptions);

            // Highlight found cells
            foreach(Cell cell in searchResult)
                cell.Fill.BackgroundColor = highlightColor;

            // Create output stream and save document in target format
            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
