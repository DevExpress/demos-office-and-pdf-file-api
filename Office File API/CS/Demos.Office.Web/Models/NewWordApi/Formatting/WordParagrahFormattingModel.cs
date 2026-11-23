using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Word;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting {
    public class WordParagraphFormattingModel : WordFileModelBase {
        [Required(ErrorMessage = "Please select an action type")]
        public ParagraphFormattingActionType ActionType { get; set; } = ParagraphFormattingActionType.ApplyFormatting;
        [Required(ErrorMessage = "Please select a range type")]
        public ParagraphFormattingRangeType RangeType { get; set; } = ParagraphFormattingRangeType.FirstParagraph;
        public bool ApplyParagraphBorders { get; set; } = false;

        public IEnumerable<ParagraphFormattingAttribute> ParagraphResetValues { get; set; }

        public ParagraphAlignment ParagraphAlignment { get; set; } = ParagraphAlignment.Left;
        public float LineSpacingMultiplier { get; set; } = 1.15f;
        public float LeftIndent { get; set; } = 0f;
        public string ShadingHex { get; set; }
        internal object Shading { get; set; } = "#FFFFFF";
        public bool WidowOrphanControl { get; set; } = true;
        public bool KeepWithNext { get; set; } = false;
        public bool KeepLinesTogether { get; set; } = false;

        public float BorderLineWidth { get; set; } = 2f;
        public BorderLineStyle BorderLineStyle { get; set; } = BorderLineStyle.Single;
        public string BorderColorHex { get; set; }
        public object BorderColor { get; set; } = "#4682B4";

        public bool IsApplyAction { get { return ActionType == ParagraphFormattingActionType.ApplyFormatting; } }
        public bool IsClearAction { get { return ActionType == ParagraphFormattingActionType.ClearFormatting; } }

        public override string OutputName { get; set; } = "Formatted_Document";
        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" },
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal readonly List<object> ActionTypes = new() {
            new { Text = "Apply Formatting", Value = ParagraphFormattingActionType.ApplyFormatting },
            new { Text = "Clear Formatting", Value = ParagraphFormattingActionType.ClearFormatting }
        };
        internal readonly List<object> RangeTypes = new() {
            new { Text = "First Paragraph", Value = ParagraphFormattingRangeType.FirstParagraph },
            new { Text = "First Section", Value = ParagraphFormattingRangeType.FirstSection },
            new { Text = "Entire Document", Value = ParagraphFormattingRangeType.WholeDocument }
        };
        internal readonly List<object> LineTypes = new() {
            new { Text = "None", Value = BorderLineStyle.None },
            new { Text = "Single", Value = BorderLineStyle.Single },
            new { Text = "Double", Value = BorderLineStyle.Double },
            new { Text = "Dotted", Value = BorderLineStyle.Dotted },
            new { Text = "Dashed", Value = BorderLineStyle.Dashed },
            new { Text = "Wavy", Value = BorderLineStyle.Wave }
        };
        internal readonly List<object> AlignmentTypes = new() {
            new { Text = "Left", Value = ParagraphAlignment.Left },
            new { Text = "Right", Value = ParagraphAlignment.Right },
            new { Text = "Center", Value = ParagraphAlignment.Center },
            new { Text = "Justify", Value = ParagraphAlignment.Justify }
        };
        internal readonly List<object> ParagraphResetSettings = new() {
            new { Text = "Alignment", Value = ParagraphFormattingAttribute.Alignment },
            new { Text = "Line Spacing", Value = ParagraphFormattingAttribute.LineSpacing },
            new { Text = "Spacing Before", Value = ParagraphFormattingAttribute.SpacingBefore },
            new { Text = "Spacing After", Value = ParagraphFormattingAttribute.SpacingAfter },
            new { Text = "Left Indent", Value = ParagraphFormattingAttribute.LeftIndent },
            new { Text = "Shading", Value = ParagraphFormattingAttribute.Shading },
            new { Text = "Widow/Orphan Control", Value = ParagraphFormattingAttribute.WidowOrphanControl },
            new { Text = "Keep with Next", Value = ParagraphFormattingAttribute.KeepWithNext },
            new { Text = "Keep Lines Together", Value = ParagraphFormattingAttribute.KeepLinesTogether }
        };
    }
    public enum ParagraphFormattingActionType {
        ApplyFormatting,
        ClearFormatting
    }
    public enum ParagraphFormattingRangeType {
        FirstParagraph,
        FirstSection,
        WholeDocument
    }
    public enum ParagraphFormattingAttribute {
        Alignment,
        LineSpacing,
        SpacingBefore,
        SpacingAfter,
        LeftIndent,
        Shading,
        WidowOrphanControl,
        KeepWithNext,
        KeepLinesTogether
    }
}