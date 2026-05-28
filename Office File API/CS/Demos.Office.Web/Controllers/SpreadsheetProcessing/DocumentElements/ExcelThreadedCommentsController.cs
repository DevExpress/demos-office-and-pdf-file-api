#region Code
using DevExpress.Spreadsheet;
#endregion Code

using Microsoft.AspNetCore.Mvc;
using BlazorDemo.OfficeFileApi.Controllers.Spreadsheet;
using BlazorDemo.OfficeFileApi.Services;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    [ApiController]
    [Route("api/spreadsheetprocessing/")]
    public class ExcelThreadedCommentsController : SpreadsheetControllerBase {
        [HttpPost]
        [HttpPost("threaded-comments/generate")]
        public async Task<IActionResult> GenerateThreadedCommentsWorkbook([FromForm] ExcelThreadedCommentsModel model) {
            // Determine output format
            var outputFormat = ExcelFormatHelper.GetDocumentFormat(model.OutputFormat);

            // Generate document with a threaded comment and a reply
            var convertedStream = GenerateDocumentWithThreadedComments(
                model.ThreadAuthor,
                model.ReplyAuthor,
                model.FirstCommentText,
                model.ReplyCommentText,
                model.Resolved,
                outputFormat);

            return CreateOutputFile(convertedStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream GenerateDocumentWithThreadedComments(string threadAuthor, string replyAuthor,
            string firstCommentText, string replyCommentText, bool resolved, DocumentFormat outputFormat) {
            using var workbook = new Workbook();
            var worksheet = workbook.Worksheets.First();

            worksheet.Cells["A1"].Value = "Select the 'B3' cell to read comments";
            worksheet.Cells["A3"].Value = "Discount:";

            var commentCell = worksheet.Cells["B3"];
            commentCell.Value = 0.15;
            commentCell.NumberFormat = "0%";
            worksheet.SelectedCell = commentCell;

            // Add a threaded comment to cell 'B3'
            var comments = worksheet.ThreadedComments;

            ThreadedComment threadedComment = comments.Add(commentCell, threadAuthor, firstCommentText);
            threadedComment.Resolved = resolved;

            // Add a reply
            threadedComment.Reply(replyAuthor, replyCommentText);

            // Prepare output stream
            var outputStream = new MemoryStream();

            // Save the document
            workbook.SaveDocument(outputStream, outputFormat);

            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}
