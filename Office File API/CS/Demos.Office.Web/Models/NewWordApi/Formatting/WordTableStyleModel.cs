using BlazorDemo.OfficeFileApi.Models.WordProcessing;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.Formatting {
    public class WordTableStyleModel : WordFileModelBase {
        [Flags]
        public enum DemoTableLookTypes {
            None = 0x0000,
            FirstRow = 0x0020,
            LastRow = 0x0040,
            FirstColumn = 0x0080,
            LastColumn = 0x0100,
            RowBanding = 0x0200,
            ColumnBanding = 0x0400
        }
        public IEnumerable<DemoTableLookTypes> DemoTableLookValues { get; set; } = new List<DemoTableLookTypes>() { DemoTableLookTypes.FirstRow };
        public int TableLookValues { get; set; }
        public override string OutputName { get; set; } = "Table_Style";
        internal readonly List<object> ConditionalTableStyleSettings = new() {
            new { Text = "Header Row", Value = DemoTableLookTypes.FirstRow },
            new { Text = "First Column", Value = DemoTableLookTypes.FirstColumn },
            new { Text = "Total Row", Value = DemoTableLookTypes.LastRow },
            new { Text = "Last Column", Value = DemoTableLookTypes.LastColumn },
            new { Text = "Column Banding", Value = DemoTableLookTypes.ColumnBanding },
            new { Text = "Row Banding", Value = DemoTableLookTypes.RowBanding }
        };
        internal readonly List<FormatInfo> OutputFormats = new() {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
    }
}