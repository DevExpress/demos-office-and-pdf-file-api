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
    public class ExcelClearContentsController : SpreadsheetControllerBase {
        [HttpPost("clear-contents/clear")]
        public async Task<IActionResult> Clear(IFormFile file, [FromForm] ExcelClearContentsModel model) {
            // Read file stream
            using var inputStream = new MemoryStream();
            await file.CopyToAsync(inputStream);

            inputStream.Seek(0, SeekOrigin.Begin);

            // Determine target format
            var targetFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Clear cell contents on the first worksheet
            var documentStream = ClearContents(inputStream, model.ClearMode, targetFormat);

            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream ClearContents(Stream inputStream, ClearMode clearMode, DocumentFormat outputFormat) {
            using Workbook workbook = new Workbook();
            workbook.LoadDocument(inputStream);

            // Get first worksheet
            Worksheet firstWorksheet = workbook.Worksheets.First();

            // Get used worksheet range
            CellRange rangeToClear = firstWorksheet.GetUsedRange();

            // Clear cell range based on the specified clear mode
            switch(clearMode) {
                case ClearMode.All:
                    ClearAll(firstWorksheet, rangeToClear);
                    break;
                case ClearMode.CellContent:
                    ClearCellContent(firstWorksheet, rangeToClear);
                    break;
                case ClearMode.CellFormatting:
                    ClearCellFormatting(firstWorksheet, rangeToClear, workbook.Styles.DefaultStyle);
                    break;
                case ClearMode.CellHyperlinks:
                    ClearHyperlinks(firstWorksheet, rangeToClear);
                    break;
                case ClearMode.CellComments:
                    ClearComments(firstWorksheet, rangeToClear);
                    break;
            }

            // Create output stream
            var outputStream = new MemoryStream();

            // Save document in target format
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Seek(0, SeekOrigin.Begin);

            return outputStream;
        }

        void ClearAll(Worksheet worksheet, CellRange rangeToClear) {
            // Remove all cell information (content, formatting, hyperlinks and comments)
            worksheet.Clear(rangeToClear);
        }

        void ClearCellContent(Worksheet worksheet, CellRange rangeToClear) {
            // Remove cell content
            worksheet.ClearContents(rangeToClear);
            rangeToClear.Value = null;
        }

        void ClearCellFormatting(Worksheet worksheet, CellRange rangeToClear, Style defaultStyle) {
            // Remove cell formatting
            worksheet.ClearFormats(rangeToClear);
            rangeToClear.Style = defaultStyle;
        }

        void ClearHyperlinks(Worksheet worksheet, CellRange rangeToClear) {
            // Remove hyperlinks from cells
            worksheet.ClearHyperlinks(rangeToClear);
            IList<Hyperlink> hyperlinks = worksheet.Hyperlinks.GetHyperlinks(rangeToClear);
            if(hyperlinks.Count > 0) {
                foreach(Hyperlink hyperlink in hyperlinks) {
                    worksheet.Hyperlinks.Remove(hyperlink);
                }
            }
        }

        void ClearComments(Worksheet worksheet, CellRange rangeToClear) {
            // Remove comments from cells.
            worksheet.ClearComments(rangeToClear);
        }
        #endregion Code
    }
}
