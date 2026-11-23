using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.ReviewComment {
    public class WordCommentsModel : WordFileModelBase, IValidatableObject {
        [Required(ErrorMessage = "Comment action is required")]
        public CommentActionType CommentAction { get; set; } = CommentActionType.AddNew;
        public RemoveCommentActionType RemoveCommentAction { get; set; } = RemoveCommentActionType.ByAuthor;

        public string AuthorFilter { get; set; } = "Author1";

        public string CommentText { get; set; } = "Test Comment";

        public string CommentAuthor { get; set; } = "TestUser";

        public string ReplyText { get; set; } = "Test Reply";

        public string ReplyAuthor { get; set; } = "TestUser";

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "Comments_Result";
        internal bool IsAddComment => CommentAction == CommentActionType.AddNew;
        internal bool IsAddReply => CommentAction == CommentActionType.Reply;
        internal bool IsRemoveComments => CommentAction == CommentActionType.Remove;
        internal bool IsRemoveCommentByAuthor => IsRemoveComments && RemoveCommentAction == RemoveCommentActionType.ByAuthor;

        internal readonly List<object> CommentActions = new() {
            new { Text = "Add New Comment", Value = CommentActionType.AddNew },
            new { Text = "Add Reply to Comment", Value = CommentActionType.Reply },
            new { Text = "Remove Comments", Value = CommentActionType.Remove }
        };
        internal readonly List<object> RemoveCommentActions = new() {
            new { Text = "By Author", Value = RemoveCommentActionType.ByAuthor },
            new { Text = "All Comments", Value = RemoveCommentActionType.AllComments }
        };

        internal readonly List<object> AuthorFilters = new() {
            new { Text = "Author1", Value = "Author1" },
            new { Text = "Author2", Value = "Author2" },
            new { Text = "Author3", Value = "Author3" }
        };

        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext) {
            var results = new List<ValidationResult>();

            switch (CommentAction) {
                case CommentActionType.AddNew:
                    if (string.IsNullOrWhiteSpace(CommentText)) {
                        results.Add(new ValidationResult(
                            "Comment text is required when adding a comment",
                            new[] { nameof(CommentText) }));
                    }
                    if (string.IsNullOrWhiteSpace(CommentAuthor)) {
                        results.Add(new ValidationResult(
                            "Comment author is required when adding a comment",
                            new[] { nameof(CommentAuthor) }));
                    }
                    break;

                case CommentActionType.Reply:
                    if (string.IsNullOrWhiteSpace(ReplyText)) {
                        results.Add(new ValidationResult(
                            "Reply text is required when adding a reply",
                            new[] { nameof(ReplyText) }));
                    }
                    if (string.IsNullOrWhiteSpace(ReplyAuthor)) {
                        results.Add(new ValidationResult(
                            "Reply author is required when adding a reply",
                            new[] { nameof(ReplyAuthor) }));
                    }
                    break;
            }

            return results;
        }
    }
    public enum CommentActionType {
        AddNew,
        Reply,
        Remove
    }
    public enum RemoveCommentActionType {
        ByAuthor,
        AllComments
    }
}