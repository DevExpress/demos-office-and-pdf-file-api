using BlazorDemo.OfficeFileApi.Models.WordProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentConversion {
    public class WordToWordConversionModel : WordFileModelBase {
        public List<FormatInfo> OutputFormats { get; set; } = new()
        {
            new FormatInfo() { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo() { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo() { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo() { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo() { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo() { Text = "Rich Text Format (.rtf)", Value = "rtf" },
            new FormatInfo() { Text = "Web Archive (.mht)", Value = "mht" },
            new FormatInfo() { Text = "Web Page (.html)", Value = "html" },
            new FormatInfo() { Text = "Plain Text (.txt)", Value = "txt" },
            new FormatInfo() { Text = "OpenDocument Text (.odt)", Value = "odt" }
        };
        public override string OutputName { get; set; } = "ConvertedDocument";
        public bool EmbedFonts { get; set; } = false;
    }
}