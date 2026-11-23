using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Office;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting {
    public class WordTextFormattingModel : WordFileModelBase {
        [Required(ErrorMessage = "Please select an action type")]
        public TextFormattingActionType ActionType { get; set; } = TextFormattingActionType.ApplyFormatting;
        [Required(ErrorMessage = "Please select a range type")]
        public TextFormattingRangeType RangeType { get; set; } = TextFormattingRangeType.FirstParagraph;

        public string FontName { get; set; } = "Segoe UI";
        public float FontSize { get; set; } = 16f;
        public TextFontStyle FontStyle { get; set; } = TextFontStyle.Regular;
        public string ForeColorHex { get; set; }
        internal object ForeColor { get; set; } = "#000000";
        public string BackColorHex { get; set; }
        internal object BackColor { get; set; } = "#FFFFFF";
        public TextUnderlineType Underline { get; set; } = TextUnderlineType.None;

        public IEnumerable<TextFormattingAttribute> CharacterResetValues { get; set; }

        public bool IsApplyAction { get { return ActionType == TextFormattingActionType.ApplyFormatting; } }
        public bool IsClearAction { get { return ActionType == TextFormattingActionType.ClearFormatting; } }
        public bool IsItalic { get { return FontStyle is TextFontStyle.Italic or TextFontStyle.BoldItalic; } }
        public bool IsBold { get { return FontStyle is TextFontStyle.Bold or TextFontStyle.BoldItalic; } }

        public override string OutputName { get; set; } = "Formatted_Document";
        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "PDF Document (.pdf)", Value = "pdf" },
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal readonly List<object> ActionTypes = new() {
            new { Text = "Apply Formatting", Value = TextFormattingActionType.ApplyFormatting },
            new { Text = "Clear Formatting", Value = TextFormattingActionType.ClearFormatting },
        };
        internal readonly List<object> RangeTypes = new() {
            new { Text = "First Paragraph", Value = TextFormattingRangeType.FirstParagraph },
            new { Text = "First Section", Value = TextFormattingRangeType.FirstSection },
            new { Text = "Entire Document", Value = TextFormattingRangeType.WholeDocument }
        };
        internal readonly List<object> FontNames = new() {
            new { Text = "Calibri", Value = "Calibri" },
            new { Text = "Segoe UI", Value = "Segoe UI" },
            new { Text = "Arial", Value = "Arial" },
            new { Text = "Times New Roman", Value = "Times New Roman" },
            new { Text = "Verdana", Value = "Verdana" },
            new { Text = "Tahoma", Value = "Tahoma" },
            new { Text = "Georgia", Value = "Georgia" },
            new { Text = "Impact", Value = "Impact" }
        };
        internal readonly List<object> UnderlineTypes = new() {
            new { Text = "None", Value = TextUnderlineType.None },
            new { Text = "Single", Value = TextUnderlineType.Single },
            new { Text = "Double", Value = TextUnderlineType.Double },
            new { Text = "Dotted", Value = TextUnderlineType.Dotted },
            new { Text = "Dashed", Value = TextUnderlineType.Dashed },
            new { Text = "Wave", Value = TextUnderlineType.Wavy }
        };
        internal readonly List<object> CharacterResetSettings = new() {
            new { Text = "Font Name", Value = TextFormattingAttribute.FontName },
            new { Text = "Font Size", Value = TextFormattingAttribute.FontSize },
            new { Text = "Font Color", Value = TextFormattingAttribute.ForeColor },
            new { Text = "Italic", Value = TextFormattingAttribute.Italic },
            new { Text = "Bold", Value = TextFormattingAttribute.Bold },
            new { Text = "Background Color", Value = TextFormattingAttribute.BackColor },
            new { Text = "Underline", Value = TextFormattingAttribute.Underline }
        };
        internal readonly List<object> FontStyles = new() {
            new { Text = "Regular", Value = TextFontStyle.Regular },
            new { Text = "Italic", Value = TextFontStyle.Italic },
            new { Text = "Bold", Value = TextFontStyle.Bold },
            new { Text = "Bold Italic", Value = TextFontStyle.BoldItalic }
        };
    }
    public enum TextFormattingActionType {
        ApplyFormatting,
        ClearFormatting
    }
    public enum TextFormattingRangeType {
        FirstParagraph,
        FirstSection,
        WholeDocument
    }
    public enum TextFontStyle {
        Regular,
        Italic,
        Bold,
        BoldItalic
    }
    public enum TextFormattingAttribute {
        FontName,
        FontSize,
        ForeColor,
        Italic,
        Bold,
        BackColor,
        Underline
    }
}