#region Code
using DevExpress.Office;
using DevExpress.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
#endregion Code
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.Spreadsheet.ManageSheets;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.SpreadsheetProcessing.ManageSheets {
    [ApiController]
    [Route("api/spreadsheetprocessing")]
    public class ExcelFormatRowsColumnsController : SpreadsheetControllerBase {
        [HttpPost("format-rows-columns/rows/add")]
        public async Task<IActionResult> AddRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Add rows to the first worksheet
            var convertedStream = AddRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        // Rows
        Stream AddRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Insert rows into the worksheet at the specified position.
            firstWorksheet.Rows.Insert(startRowIndex, rowCount);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/remove")]
        public async Task<IActionResult> RemoveRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Remove rows from the first worksheet
            var convertedStream = RemoveRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream RemoveRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Remove rows from the worksheet starting with the specified position.
            firstWorksheet.Rows.Remove(startRowIndex, rowCount);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/setsize")]
        public async Task<IActionResult> SetSizeRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentUnit = GetDocumentUnit(model.DocumentUnit);

            // Set rows height
            var convertedStream = SetSizeRows(inputStream, model.RowCount, model.StartRowIndex - 1, documentUnit, model.RowHeight, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SetSizeRows(Stream inputStream, int rowCount, int startRowIndex, DocumentUnit unit, double rowHeight, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            workbook.Unit = unit;

            // Set rows height
            for (int i = startRowIndex; i < startRowIndex + rowCount; i++) {
                firstWorksheet.Rows[i].Height = rowHeight;
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/hide")]
        public async Task<IActionResult> HideRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Hide specified rows
            var convertedStream = HideRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream HideRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Hide rows starting with the specified position.
            firstWorksheet.Rows.Hide(startRowIndex, startRowIndex + rowCount - 1);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/autofit")]
        public async Task<IActionResult> AutoFitRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Apply Auto Fit to the specified rows
            var convertedStream = AutoFitRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AutoFitRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Auto Fit rows starting with the specified position
            firstWorksheet.Rows.AutoFit(startRowIndex, startRowIndex + rowCount - 1);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/add")]
        public async Task<IActionResult> AddColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Add columns to the first worksheet
            var convertedStream = AddColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        // Columns
        Stream AddColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Insert columns into the worksheet at the specified position
            firstWorksheet.Columns.Insert(startColumnIndex, columnCount);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/remove")]
        public async Task<IActionResult> RemoveColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Remove columns from the first worksheet
            var convertedStream = RemoveColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream RemoveColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Remove columns from the worksheet starting with the specified position
            firstWorksheet.Columns.Remove(startColumnIndex, columnCount);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/setsize")]
        public async Task<IActionResult> SetSizeColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentUnit = GetDocumentUnit(model.DocumentUnit);

            // Set columns width
            var convertedStream = SetSizeColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, documentUnit, model.ColumnWidth, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SetSizeColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentUnit unit, double columnWidth, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            workbook.Unit = unit;

            // Set columns height
            for (int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
                firstWorksheet.Columns[i].Width = columnWidth;
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/hide")]
        public async Task<IActionResult> HideColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Hide specified columns
            var convertedStream = HideColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream HideColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Hide columns starting with the specified position
            firstWorksheet.Columns.Hide(startColumnIndex, startColumnIndex + columnCount - 1);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/autofit")]
        public async Task<IActionResult> AutoFitColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Apply Auto Fit to the specified columns
            var convertedStream = AutoFitColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AutoFitColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Auto Fit columns starting with the specified position
            firstWorksheet.Columns.AutoFit(startColumnIndex, startColumnIndex + columnCount - 1);

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }
        #endregion Code
        DocumentUnit GetDocumentUnit(string unit) {
            return unit.ToLowerInvariant() switch {
                "centimeter" => DocumentUnit.Centimeter,
                "document" => DocumentUnit.Document,
                "inch" => DocumentUnit.Inch,
                "millimeter" => DocumentUnit.Millimeter,
                "point" => DocumentUnit.Point,
                _ => DocumentUnit.Document
            };
        }
    }
}
