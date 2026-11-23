#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.ManageSheets {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelWorksheetOperationsController : SpreadsheetControllerBase {
        [HttpPost("worksheet-operations/add")]
        public async Task<IActionResult> Add(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = AddWorksheet(inputStream, model.WorksheetName, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddWorksheet(Stream inputStream, string worksheetName, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            if(workbook.Worksheets.Contains(worksheetName)) {
                worksheetName += "_New";
            }

            workbook.Worksheets.Add(new Worksheet(worksheetName));

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code

        [HttpPost("worksheet-operations/remove")]
        public async Task<IActionResult> Remove(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var worksheetRange = RangeHelper.ParseRange(model.WorksheetRange);

            var convertedStream = RemoveWorksheet(inputStream, worksheetRange.ToArray(), targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream RemoveWorksheet(Stream inputStream, int[] worksheetIndexes, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            List<Worksheet> worksheetsToRemove = new List<Worksheet>();
            int worksheetCount = workbook.Worksheets.Count;
            foreach(int index in worksheetIndexes) {
                if(index < worksheetCount) {
                    worksheetsToRemove.Add(workbook.Worksheets[index]);
                }
            }

            foreach(var worksheetToRemove in worksheetsToRemove) {
                try {
                    workbook.Worksheets.Remove(worksheetToRemove);
                }
                catch(InvalidOperationException) {
                    continue;
                }
            }

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code

        [HttpPost("worksheet-operations/duplicate")]
        public async Task<IActionResult> Duplicate(IFormFile file, [FromForm] ExcelWorksheetOperationsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var worksheetRange = RangeHelper.ParseRange(model.WorksheetRange);

            var convertedStream = DuplicateWorksheet(inputStream, worksheetRange.ToArray(), targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream DuplicateWorksheet(Stream inputStream, int[] worksheetIndexes, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            List<Worksheet> worksheetsToDuplicate = new List<Worksheet>();
            int worksheetCount = workbook.Worksheets.Count;
            foreach(int index in worksheetIndexes) {
                if(index < worksheetCount) {
                    worksheetsToDuplicate.Add(workbook.Worksheets[index]);
                }
            }

            foreach(var sourceWorksheet in worksheetsToDuplicate) {
                Worksheet newWorksheet = sourceWorksheet.Clone();
                newWorksheet.Name = $"{sourceWorksheet.Name}_Copy";

                workbook.Worksheets.Add(newWorksheet);
            }

            var outputStream = new MemoryStream();
            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
