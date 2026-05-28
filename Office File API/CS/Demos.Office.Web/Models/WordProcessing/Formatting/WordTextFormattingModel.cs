using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.Formatting
{
    public class WordTextFormattingModel : WordFileModelBase
    {
        [Required(ErrorMessage = "Please select an action type")]
        public FormattingActionType ActionType { get; set; } = FormattingActionType.ApplyFormatting;
        [Required(ErrorMessage = "Please select a range type")]
        public FormattedRangeType RangeType { get; set; } = FormattedRangeType.FirstParagraph;

        // ----- Character formatting (Apply mode) -----
        public string FontName { get; set; } = "Segoe UI";
        public float FontSize { get; set; } = 16f;
        public FontStyle FontStyle { get; set; } = FontStyle.Regular;
        public string ForeColorHex { get; set; }
        internal object ForeColor { get; set; } = "#000000";
        public string BackColorHex { get; set; }
        internal object BackColor { get; set; } = "#FFFFFF";
        public UnderlineType Underline { get; set; } = UnderlineType.None;

        // ----- Reset flags (Clear mode) -----
        public IEnumerable<CharacterPropertiesMask> CharacterResetValues { get; set; }

        public bool IsApplyAction { get { return ActionType == FormattingActionType.ApplyFormatting; } }
        public bool IsClearAction { get { return ActionType == FormattingActionType.ClearFormatting; } }
        public bool HasUnderline { get { return Underline == UnderlineType.None; } }
        public bool IsItalic { get { return FontStyle == FontStyle.Italic || FontStyle == FontStyle.BoldItalic; } }
        public bool IsBold { get { return FontStyle == FontStyle.Bold || FontStyle == FontStyle.BoldItalic; } }

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
            new { Text = "Clear Formatting", Value = FormattingActionType.ClearFormatting },
        };
        internal readonly List<object> RangeTypes = new()
        {
            new { Text = "First Paragraph", Value = FormattedRangeType.FirstParagraph },
            new { Text = "First Page", Value = FormattedRangeType.FirstPage },
            new { Text = "Entire Document", Value = FormattedRangeType.WholeDocument }
        };
        internal readonly List<object> FontNames = new()
        {
            new { Text = "Calibri", Value = "Calibri" },
            new { Text = "Segoe UI", Value = "Segoe UI" },
            new { Text = "Arial", Value = "Arial" },
            new { Text = "Times New Roman", Value = "Times New Roman" },
            new { Text = "Verdana", Value = "Verdana" },
            new { Text = "Tahoma", Value = "Tahoma" },
            new { Text = "Georgia", Value = "Georgia" },
            new { Text = "Impact", Value = "Impact" }
        };
        internal readonly List<object> UnderlineTypes = new()
        {
            new { Text = "None", Value = UnderlineType.None },
            new { Text = "Single", Value = UnderlineType.Single },
            new { Text = "Double", Value = UnderlineType.Double },
            new { Text = "Dotted", Value = UnderlineType.Dotted },
            new { Text = "Dashed", Value = UnderlineType.Dashed },
            new { Text = "Wave", Value = UnderlineType.Wave }
        };

        internal readonly List<object> CharacterResetSettings = new()
        {
            new { Text = "Font Name", Value = CharacterPropertiesMask.FontName | CharacterPropertiesMask.FontNameAscii | CharacterPropertiesMask.FontNameComplexScript | CharacterPropertiesMask.FontNameEastAsia | CharacterPropertiesMask.FontNameHighAnsi },
            new { Text = "Font Size", Value = CharacterPropertiesMask.FontSize },
            new { Text = "Font Color", Value = CharacterPropertiesMask.ForeColor },
            new { Text = "Italic", Value = CharacterPropertiesMask.Italic },
            new { Text = "Bold", Value = CharacterPropertiesMask.Bold },
            new { Text = "Background Color", Value = CharacterPropertiesMask.BackColor },
            new { Text = "Underline", Value = CharacterPropertiesMask.UnderlineColor | CharacterPropertiesMask.UnderlineType }
        };
        internal readonly List<object> FontStyles = new()
        {
            new { Text = "Regular", Value = FontStyle.Regular },
            new { Text = "Italic", Value = FontStyle.Italic },
            new { Text = "Bold", Value = FontStyle.Bold },
            new { Text = "Bold Italic", Value = FontStyle.BoldItalic }
        };
    }
    public enum FontStyle
    {
        Regular,
        Italic,
        Bold,
        BoldItalic
    }
}
