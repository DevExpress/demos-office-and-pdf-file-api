using DevExpress.XtraRichEdit.API.Native;
using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentElements
{
    public class WordFootnotesEndnotesModel : WordFileModelBase
    {
        [Display(Name = "Add New Footnote")]
        public bool AddNewFootnote { get; set; } = false;

        [Display(Name = "Add New Endnote")]
        public bool AddNewEndnote { get; set; } = false;

        [Display(Name = "Note Separator")]
        public string NoteSeparator { get; set; } = " - ";

        [Required(ErrorMessage = "Note restart type is required")]
        [Display(Name = "Note Restart Type")]
        public NoteRestartType NoteRestartType { get; set; } = NoteRestartType.Continuous;

        [Required(ErrorMessage = "Numbering format is required")]
        [Display(Name = "Numbering Format")]
        public NumberingFormat NumberingFormat { get; set; } = NumberingFormat.Decimal;

        [Range(1, 1000, ErrorMessage = "Start number must be between 1 and 1000")]
        [Display(Name = "Start Number")]
        public int StartNumber { get; set; } = 1;

        [Required(ErrorMessage = "Endnote position is required")]
        [Display(Name = "Endnote Position")]
        public EndnotePosition EndnotePosition { get; set; } = EndnotePosition.EndOfDocument;

        [Required(ErrorMessage = "Footnote position is required")]
        [Display(Name = "Footnote Position")]
        public FootnotePosition FootnotePosition { get; set; } = FootnotePosition.BottomOfPage;
        internal bool IsContinuesNumbering => NoteRestartType == NoteRestartType.Continuous;

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" }
        };

        internal readonly List<object> NoteRestartTypes = new()
        {
            new { Text = "Continuous", Value = NoteRestartType.Continuous },
            new { Text = "New Page", Value = NoteRestartType.NewPage },
            new { Text = "New Section", Value = NoteRestartType.NewSection }
        };

        internal readonly List<object> NumberingFormats = new()
        {
            new { Text = "1, 2, 3... (Decimal)", Value = NumberingFormat.Decimal },
            new { Text = "I, II, III... (Roman Upper)", Value = NumberingFormat.UpperRoman },
            new { Text = "i, ii, iii... (Roman Lower)", Value = NumberingFormat.LowerRoman },
            new { Text = "A, B, C... (Letter Upper)", Value = NumberingFormat.UpperLetter },
            new { Text = "a, b, c... (Letter Lower)", Value = NumberingFormat.LowerLetter }
        };

        internal readonly List<object> EndnotePositions = new()
        {
            new { Text = "End of Document", Value = EndnotePosition.EndOfDocument },
            new { Text = "End of Section", Value = EndnotePosition.EndOfSection }
        };

        internal readonly List<object> FootnotePositions = new()
        {
            new { Text = "Bottom of Page", Value = FootnotePosition.BottomOfPage },
            new { Text = "Below Text", Value = FootnotePosition.BelowText }
        };

        public override List<string> AllowedFileExtensions { get; set; } = new List<string> { ".docx", ".docm", ".dotx", ".dotm" };
        public override string OutputName { get; set; } = "FootnotesEndnotes_Result";
    }
}
