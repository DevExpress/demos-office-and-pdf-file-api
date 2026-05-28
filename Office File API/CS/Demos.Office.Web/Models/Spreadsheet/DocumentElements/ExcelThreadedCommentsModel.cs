using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.Spreadsheet.DocumentElements {
    public class ExcelThreadedCommentsModel : ExcelFileModelBase {
        public override string OutputName { get; set; } = "ThreadedComments_Document";

        [Required(ErrorMessage = "Thread Author is required")]
        [MaxLength(40, ErrorMessage = "Thread Author must not exceed 40 characters")]
        public string ThreadAuthor { get; set; } = "Sales Department";

        [Required(ErrorMessage = "First Comment Text is required")]
        [MaxLength(120, ErrorMessage = "First Comment Text must not exceed 120 characters")]
        public string FirstCommentText { get; set; } = "The discount was approved by the Sales Director.";

        [Required(ErrorMessage = "Reply Author is required")]
        [MaxLength(40, ErrorMessage = "Reply Author must not exceed 40 characters")]
        public string ReplyAuthor { get; set; } = "Jane Doe";

        [Required(ErrorMessage = "Reply Text is required")]
        [MaxLength(120, ErrorMessage = "Reply Text must not exceed 120 characters")]
        public string ReplyCommentText { get; set; } = "Thank you for the confirmation!";

        public bool Resolved { get; set; } = true;
    }
}
