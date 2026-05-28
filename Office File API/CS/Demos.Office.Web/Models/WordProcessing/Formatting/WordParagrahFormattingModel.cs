using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.Formatting
{
    public class WordParagraphFormattingModel : WordFileModelBase
    {
        [Required(ErrorMessage = "Please select an action type")]
        public FormattingActionType ActionType { get; set; } = FormattingActionType.ApplyFormatting;
        [Required(ErrorMessage = "Please select a range type")]
        public FormattedRangeType RangeType { get; set; } = FormattedRangeType.FirstParagraph;
        public bool ApplyParagraphBorders { get; set; } = false;


        // ----- Reset flags (Clear mode) -----
        public IEnumerable<ParagraphPropertiesMask> ParagraphResetValues { get; set; }

        // ----- Paragraph formatting (Apply mode) -----
        public ParagraphAlignment ParagraphAlignment { get; set; } = ParagraphAlignment.Left;
        public float LineSpacingMultiplier { get; set; } = 1.15f;
        public float LeftIndent { get; set; } = 0f;
        public string ShadingHex { get; set; }
        internal object Shading { get; set; } = "#FFFFFF";
        public bool WidowOrphanControl { get; set; } = true;
        public bool KeepWithNext { get; set; } = false;
        public bool KeepLinesTogether { get; set; } = false;


        // Paragraph borders (Apply mode)
        public float BorderLineWidth { get; set; } = 2f;
        public BorderLineStyle BorderLineStyle { get; set; } = BorderLineStyle.Single;
        public string BorderColorHex { get; set; }
        public object BorderColor { get; set; } = "#4682B4";

        public bool IsApplyAction { get { return ActionType == FormattingActionType.ApplyFormatting; } }
        public bool IsClearAction { get { return ActionType == FormattingActionType.ClearFormatting; } }

        public override string OutputName { get; set; } = "Formatted_Document";
        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" },
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal readonly List<object> ActionTypes = new()
        {
            new { Text = "Apply Formatting", Value = FormattingActionType.ApplyFormatting },
            new { Text = "Clear Formatting", Value = FormattingActionType.ClearFormatting }
        };
        internal readonly List<object> RangeTypes = new()
        {
            new { Text = "First Paragraph", Value = FormattedRangeType.FirstParagraph },
            new { Text = "First Page", Value = FormattedRangeType.FirstPage },
            new { Text = "Entire Document", Value = FormattedRangeType.WholeDocument }
        };
        internal readonly List<object> LineTypes = new()
        {
            new { Text = "None", Value = BorderLineStyle.None },
            new { Text = "Single", Value = BorderLineStyle.Single },
            new { Text = "Double", Value = BorderLineStyle.Double },
            new { Text = "Dotted", Value = BorderLineStyle.Dotted },
            new { Text = "Dashed", Value = BorderLineStyle.Dashed },
            new { Text = "Wavy", Value = BorderLineStyle.Wave }
        };
        internal readonly List<object> AlignmentTypes = new()
        {
            new { Text = "Left", Value = ParagraphAlignment.Left },
            new { Text = "Right", Value = ParagraphAlignment.Right },
            new { Text = "Center", Value = ParagraphAlignment.Center },
            new { Text = "Justify", Value = ParagraphAlignment.Justify }
        };
        internal readonly List<object> ParagraphResetSettings = new()
        {
            new { Text = "Alignment", Value = ParagraphPropertiesMask.Alignment },
            new { Text = "Line Spacing", Value = ParagraphPropertiesMask.LineSpacing },
            new { Text = "Spacing Before", Value = ParagraphPropertiesMask.SpacingBefore },
            new { Text = "Spacing After", Value = ParagraphPropertiesMask.SpacingAfter },
            new { Text = "Left Indent", Value = ParagraphPropertiesMask.LeftIndent },
            new { Text = "Shading", Value = ParagraphPropertiesMask.BackColor },
            new { Text = "Widow/Orphan Control", Value = ParagraphPropertiesMask.WidowOrphanControl },
            new { Text = "Keep with Next", Value = ParagraphPropertiesMask.KeepWithNext },
            new { Text = "Keep Lines Together", Value = ParagraphPropertiesMask.KeepLinesTogether }
        };
    }
    public enum FormattingActionType
    {
        ApplyFormatting,
        ClearFormatting
    }
    public enum FormattedRangeType
    {
        FirstParagraph,
        FirstPage,
        WholeDocument
    }
}
