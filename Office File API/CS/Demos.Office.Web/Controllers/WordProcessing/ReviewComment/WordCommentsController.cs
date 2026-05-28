using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Models.WordProcessing.ReviewComment;
using Microsoft.AspNetCore.Mvc;
#region Code
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.API.Native;
#endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.WordProcessing.ReviewComment {
    [ApiController]
    [Route("api/wordprocessing/comments")]
    public class WordCommentsController : WordControllerBase {

        [HttpGet("template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentCommentTemplate, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "Comments_Template.docx");
        }

        [HttpPost("process")]
        public async Task<IActionResult> ProcessCommentsTemplate([FromForm] WordCommentsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentCommentTemplate, FileMode.Open, FileAccess.Read);
            MemoryStream inputStream = new();
            templateStream.CopyTo(inputStream);
            inputStream.Position = 0;

            // Determine target format
            var targetFormat = WordFormatHelper.GetDocumentFormat(model.OutputFormat);
            Stream outputStream = new MemoryStream();
            switch(model.CommentAction) {
                case CommentActionType.AddNew:
                    outputStream = AddNewComment(inputStream, model.CommentText!, model.CommentAuthor!, targetFormat);
                    break;

                case CommentActionType.Reply:
                    outputStream = ReplyToComment(inputStream, model.ReplyText!, model.ReplyAuthor!, targetFormat);
                    break;

                case CommentActionType.Remove:
                    outputStream = RemoveComments(inputStream, model.RemoveCommentAction == RemoveCommentActionType.AllComments, model.AuthorFilter, targetFormat);
                    break;
            }

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddNewComment(Stream inputStream, string commentText, string author, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            Paragraph paragraph = document.Paragraphs[0];
            Comment comment = document.Comments.Create(paragraph.Range, author);
            SubDocument commentDoc = comment.BeginUpdate();
            commentDoc.InsertText(commentDoc.Range.Start, commentText);
            comment.EndUpdate(commentDoc);

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream ReplyToComment(Stream inputStream, string replyText, string author, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            // Reply to the first comment
            if(document.Comments.Count > 0) {
                Comment parentComment = document.Comments[0];
                Comment replyComment = document.Comments.Create(author, parentComment);
                SubDocument replyDoc = replyComment.BeginUpdate();
                replyDoc.InsertText(replyDoc.Range.Start, replyText);
                replyComment.EndUpdate(replyDoc);
            }

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream RemoveComments(Stream inputStream, bool removeAll, string authorFilter, DocumentFormat outputFormat) {
            using var wordProcessor = new RichEditDocumentServer();

            wordProcessor.LoadDocument(inputStream);
            var document = wordProcessor.Document;

            if(removeAll) {
                for(int i = document.Comments.Count - 1; i >= 0; i--)
                    document.Comments.Remove(document.Comments[i]);
            } else if(!string.IsNullOrWhiteSpace(authorFilter)) {
                for(int i = document.Comments.Count - 1; i >= 0; i--) {
                    if(string.Equals(document.Comments[i].Author, authorFilter, StringComparison.OrdinalIgnoreCase)) {
                        document.Comments.Remove(document.Comments[i]);
                    }
                }
            }

            var outputStream = new MemoryStream();
            wordProcessor.SaveDocument(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        #endregion Code
    }
}
