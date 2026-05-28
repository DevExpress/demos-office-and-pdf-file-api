using System.ComponentModel.DataAnnotations;
using DevExpress.Blazor.Internal;
using DevExpress.Drawing;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.ReviewComment
{
    public class WordTrackChangesModel : WordFileModelBase, IValidatableObject
    {
        [Required(ErrorMessage = "Revision action is required")]
        public RevisionActionType RevisionAction { get; set; } = RevisionActionType.Accept;

        public bool DisableTrackedChanges { get; set; } = false;

        public string AuthorFilter { get; set; } = "";

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".doc", ".docx", ".docm", ".dotx", ".dotm", ".rtf" };
        public override string OutputName { get; set; } = "TrackChanges_Result";
        internal bool IsAcceptRejectChanges => RevisionAction != RevisionActionType.None;

        internal readonly List<object> RevisionActions = new()
    {
        new { Text = "Accept All Revisions by Author", Value = RevisionActionType.Accept },
        new { Text = "Reject All Revisions by Author", Value = RevisionActionType.Reject },
        new { Text = "Keep Changes As-Is", Value = RevisionActionType.None }
    };

        internal readonly List<object> AuthorFilters = new()
    {
        new { Text = "All Authors", Value = "" },
        new { Text = "Author1", Value = "Author1" },
        new { Text = "Author2", Value = "Author2" }
    };

    internal readonly List<FormatInfo> OutputFormats = new()
    {
        new FormatInfo() { Text = "Word Document (.docx)", Value = "docx" },
        new FormatInfo() { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
        new FormatInfo() { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
        new FormatInfo() { Text = "Word Template (.dotx)", Value = "dotx" },
        new FormatInfo() { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
        new FormatInfo() { Text = "Rich Text Format (.rtf)", Value = "rtf" }
    };

        // Implement IValidatableObject for custom validation
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            var results = new List<ValidationResult>();

            // Validate revision action
            if (!new[] { RevisionActionType.None, RevisionActionType.Accept, RevisionActionType.Reject }.Contains(RevisionAction))
            {
                results.Add(new ValidationResult(
                    "Invalid revision action. Must be 'none', 'accept', or 'reject'",
                    new[] { nameof(RevisionAction) }));
            }

            return results;
        }
    }
    public enum RevisionActionType
    {
        None,
        Accept,
        Reject
    }
}
