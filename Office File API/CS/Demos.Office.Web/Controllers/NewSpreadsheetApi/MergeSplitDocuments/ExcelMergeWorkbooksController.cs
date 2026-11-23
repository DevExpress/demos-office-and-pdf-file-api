#region Code
using DevExpress.Docs.Spreadsheet;
#endregion Code
using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Models.NewSpreadsheetApi.MergeSplitDocuments;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Controllers.NewSpreadsheetApi.MergeSplitDocuments {
    [ApiController]
    [Route("api/newspreadsheetapi")]
    public class ExcelMergeWorkbooksController : SpreadsheetControllerBase {
        [HttpPost("merge-workbooks/merge")]
        public async Task<IActionResult> Merge(IFormFile[] files, [FromForm] ExcelMergeWorkbooksModel model) {
            if(files == null || files.Length == 0) {
                return BadRequest("No files are uploaded");
            }

            var inputStreams = new List<Stream>();
            foreach(var file in files) {
                var ms = new MemoryStream();
                await file.CopyToAsync(ms);
                ms.Position = 0;
                inputStreams.Add(ms);
            }

            DocumentFormat outputFormat = NewSpreadsheetFormatHelper.GetDocumentFormat(model.OutputFormat);

            var documentStream = MergeWorkbooks(inputStreams, outputFormat);

            return CreateOutputFile(documentStream, model.OutputFormat, model.OutputName);
        }
        #region Code
        Stream MergeWorkbooks(List<Stream> inputStreams, DocumentFormat outputFormat) {
            using var finalWorkbook = new Workbook(inputStreams[0]);

            for(int i = 1; i < inputStreams.Count; i++) {
                using var currentWorkbook = new Workbook(inputStreams[i]);
                finalWorkbook.Append(currentWorkbook);
            }

            var outputStream = new MemoryStream();

            finalWorkbook.Save(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
