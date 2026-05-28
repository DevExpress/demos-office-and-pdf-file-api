#region Code
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
#endregion Code

using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelWorksheetOperationsController : SpreadsheetControllerBase {
        [HttpPost("worksheet-operations/add")]
        public async Task<IActionResult> Add(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Convert document to the target format
            var convertedStream = AddWorksheet(inputStream, model.WorksheetName, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddWorksheet(Stream inputStream, string worksheetName, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Add new worksheet
            if(workbook.Worksheets.Contains(worksheetName)) {
                worksheetName += "_New";
            }
            workbook.Worksheets.Add(worksheetName);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("worksheet-operations/remove")]
        public async Task<IActionResult> Remove(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Get worksheet range
            var worksheetRange = RangeHelper.ParseRange(model.WorksheetRange);

            // Convert document to the target format
            var convertedStream = RemoveWorksheet(inputStream, worksheetRange.ToArray(), targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream RemoveWorksheet(Stream inputStream, int[] worksheetIndexes, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Get worksheets to remove
            List<Worksheet> worksheetsToRemove = new List<Worksheet>();
            int worksheetCount = workbook.Worksheets.Count;
            foreach(int index in worksheetIndexes) {
                if(index < worksheetCount) {
                    worksheetsToRemove.Add(workbook.Worksheets[index]);
                }
            }

            // Remove worksheets
            foreach(var worksheetToRemove in worksheetsToRemove) {
                try {
                    // A workbook must always contain at least one visible worksheet
                    workbook.Worksheets.Remove(worksheetToRemove);
                }
                catch {
                    continue;
                }
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code


        [HttpPost("worksheet-operations/duplicate")]
        public async Task<IActionResult> Duplicate(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Get worksheet range
            var worksheetRange = RangeHelper.ParseRange(model.WorksheetRange);

            // Convert document to the target format
            var convertedStream = DuplicateWorksheet(inputStream, worksheetRange.ToArray(), targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream DuplicateWorksheet(Stream inputStream, int[] worksheetIndexes, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Get worksheets to duplicate
            List<Worksheet> worksheetsToDuplicate = new List<Worksheet>();
            int worksheetCount = workbook.Worksheets.Count;
            foreach(int index in worksheetIndexes) {
                if(index < worksheetCount) {
                    worksheetsToDuplicate.Add(workbook.Worksheets[index]);
                }
            }

            // Duplicate worksheets
            foreach(var sourceWorksheet in worksheetsToDuplicate) {
                // Add a new worksheet to the workbook
                Worksheet newWorksheet = workbook.Worksheets.Add($"{sourceWorksheet.Name}_Copy");

                // Copy all information from source worksheet to the newly created worksheet
                newWorksheet.CopyFrom(sourceWorksheet);
            }
            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
    }
}
