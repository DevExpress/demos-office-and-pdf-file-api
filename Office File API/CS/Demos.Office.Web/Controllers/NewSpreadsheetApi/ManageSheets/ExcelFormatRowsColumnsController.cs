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
    public class ExcelFormatRowsColumnsController : SpreadsheetControllerBase {
        [HttpPost("format-rows-columns/rows/add")]
        public async Task<IActionResult> AddRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = AddRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            firstWorksheet.InsertRows(startRowIndex, rowCount);

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/remove")]
        public async Task<IActionResult> RemoveRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = RemoveRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream RemoveRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            firstWorksheet.DeleteRows(startRowIndex, rowCount);

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/setsize")]
        public async Task<IActionResult> SetSizeRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = SetSizeRows(inputStream, model.RowCount, model.StartRowIndex - 1, model.RowHeight, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SetSizeRows(Stream inputStream, int rowCount, int startRowIndex, double rowHeight, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
                firstWorksheet.Rows[i].Height = rowHeight;
            }

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/hide")]
        public async Task<IActionResult> HideRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = HideRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream HideRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
                firstWorksheet.Rows[i].Visible = false;
            }

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/rows/autofit")]
        public async Task<IActionResult> AutoFitRows(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = AutoFitRows(inputStream, model.RowCount, model.StartRowIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AutoFitRows(Stream inputStream, int rowCount, int startRowIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startRowIndex; i < startRowIndex + rowCount; i++) {
                firstWorksheet.Rows[i].AutoFit();
            }

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/add")]
        public async Task<IActionResult> AddColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = AddColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AddColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            firstWorksheet.InsertColumns(startColumnIndex, columnCount);

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/remove")]
        public async Task<IActionResult> RemoveColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = RemoveColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream RemoveColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            firstWorksheet.DeleteColumns(startColumnIndex, columnCount);

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/setsize")]
        public async Task<IActionResult> SetSizeColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = SetSizeColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, model.ColumnWidth, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream SetSizeColumns(Stream inputStream, int columnCount, int startColumnIndex, double columnWidth, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
                firstWorksheet.Columns[i].Width = columnWidth;
            }

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/hide")]
        public async Task<IActionResult> HideColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = HideColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream HideColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
                firstWorksheet.Columns[i].Visible = false;
            }

            return CreateOutputStream(workbook, outputFormat);
        }
        #endregion Code

        [HttpPost("format-rows-columns/columns/autofit")]
        public async Task<IActionResult> AutoFitColumns(IFormFile file, [FromForm] ExcelFormatRowsColumnsModel model) {
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);
            inputStream.Position = 0;

            var targetFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var convertedStream = AutoFitColumns(inputStream, model.ColumnCount, model.StartColumnIndex - 1, targetFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream AutoFitColumns(Stream inputStream, int columnCount, int startColumnIndex, DocumentFormat outputFormat) {
            using var workbook = new Workbook(inputStream);

            Worksheet firstWorksheet = workbook.Worksheets[0];

            for(int i = startColumnIndex; i < startColumnIndex + columnCount; i++) {
                firstWorksheet.Columns[i].AutoFit();
            }

            return CreateOutputStream(workbook, outputFormat);
        }

        Stream CreateOutputStream(Workbook workbook, DocumentFormat outputFormat) {
            var outputStream = new MemoryStream();

            workbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
