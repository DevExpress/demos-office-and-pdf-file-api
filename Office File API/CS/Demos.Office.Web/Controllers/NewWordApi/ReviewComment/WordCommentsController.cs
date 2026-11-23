using BlazorDemo.DemoData;
using BlazorDemo.OfficeFileApi.Controllers.WordProcessing;
using BlazorDemo.OfficeFileApi.Models.NewWordApi.ReviewComment;
using BlazorDemo.OfficeFileApi.Services;
using Microsoft.AspNetCore.Mvc;

#region Code

using DevExpress.Docs.Word;
        #endregion Code

namespace BlazorDemo.OfficeFileApi.Controllers.NewWordApi.ReviewComment {
    [ApiController]
    [Route("api/newwordapi")]
    public class WordCommentsController : WordControllerBase {
        [HttpGet("comments/template")]
        public IActionResult DownloadTemplate() {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentCommentTemplate, FileMode.Open, FileAccess.Read);
            var memoryStream = new MemoryStream();
            templateStream.CopyTo(memoryStream);
            memoryStream.Position = 0;

            return File(memoryStream.ToArray(),
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "Comments_Template.docx");
        }

        [HttpPost("comments/process")]
        public IActionResult ProcessCommentsTemplate([FromForm] WordCommentsModel model) {
            using var templateStream = new FileStream(DemoDocumentPaths.WordDocumentCommentTemplate, FileMode.Open, FileAccess.Read);
            // Determine target format
            var targetFormat = NewWordFormatHelper.GetDocumentFormat(model.OutputFormat);
            var outputStream = Stream.Null;
            switch(model.CommentAction) {
                case CommentActionType.AddNew:
                    outputStream = AddNewComment(templateStream, model.CommentText, model.CommentAuthor, targetFormat);
                    break;

                case CommentActionType.Reply:
                    outputStream = ReplyToComment(templateStream, model.ReplyText, model.ReplyAuthor, targetFormat);
                    break;

                case CommentActionType.Remove:
                    outputStream = RemoveComments(templateStream, model.RemoveCommentAction == RemoveCommentActionType.AllComments, model.AuthorFilter, targetFormat);
                    break;
            }

            return CreateOutputFile(outputStream, model.OutputFormat, model.OutputName);
        }

        #region Code
        Stream AddNewComment(Stream inputStream, string commentText, string author, DocumentFormat outputFormat) {
            using var document = new WordDocument(inputStream);

            var paragraph = document.Body.OfType<Paragraph>().First();
            var comment = document.Comments.Create(document.Body.GetRange(paragraph), new CommentDefinition { Author = author });
            comment.Content.AddText(commentText);

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream ReplyToComment(Stream inputStream, string replyText, string author, DocumentFormat outputFormat) {
            using var document = new WordDocument(inputStream);

            // Reply to the first comment
            if(document.Comments.Count > 0) {
                var parentComment = document.Comments[0];
                var reply = parentComment.CreateReply(new CommentDefinition { Author = author });
                reply.Content.AddText(replyText);
            }

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }

        Stream RemoveComments(Stream inputStream, bool removeAll, string authorFilter, DocumentFormat outputFormat) {
            using var document = new WordDocument(inputStream);

            if(removeAll) {
                for(var i = document.Comments.Count - 1; i >= 0; i--)
                    document.Comments.RemoveAt(i);
            }
            else if(!string.IsNullOrWhiteSpace(authorFilter)) {
                for(var i = document.Comments.Count - 1; i >= 0; i--) {
                    if(string.Equals(document.Comments[i].Author, authorFilter, StringComparison.OrdinalIgnoreCase)) {
                        document.Comments.RemoveAt(i);
                    }
                }
            }

            var outputStream = new MemoryStream();
            document.Save(outputStream, outputFormat);
            outputStream.Position = 0;
            return outputStream;
        }
        #endregion Code
    }
}