namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.MailMergeFields {
    public class WordFieldsModel : WordFileModelBase {
        public override string OutputName { get; set; } = "Fields_Result";
        public FieldAction Action { get; set; } = FieldAction.Insert;
        public IEnumerable<string> FieldEntryValues { get; set; }

        internal readonly List<FormatInfo> OutputFormats = new()
        {
            new FormatInfo { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
            new FormatInfo { Text = "Rich Text Format (.rtf)", Value = "rtf" }
        };
        internal readonly List<object> FieldActions = new()
        {
            new { Text = "Insert New Fields", Value = FieldAction.Insert },
            new { Text = "Delete Fields", Value = FieldAction.Delete },
            new { Text = "Lock Field Updates", Value = FieldAction.Lock },
            new { Text = "Replace Fields with Values", Value = FieldAction.Unlink },
            new { Text = "Update Field Values", Value = FieldAction.Update }
        };
        internal readonly List<object> FieldEntryTypes = new()
        {
            new { Text = "Date", Value = "DATE" },
            new { Text = "Time", Value = "TIME" },
            new { Text = "Page", Value = "PAGE" },
            new { Text = "NumPages", Value = "NUMPAGES" },
            new { Text = "DocVariable (Dynamic Field)", Value = "DocVariable" },
            new { Text = "Author", Value = "AUTHOR" },
            new { Text = "Title", Value = "TITLE" },
            new { Text = "DocProperty (Custom Field)", Value = "DOCPROPERTY" }
        };
        internal bool CanSelectFields { get { return Action != FieldAction.Update; } }
    }
    public enum FieldAction {
        Insert,
        Update,
        Lock,
        Unlink,
        Delete
    }
    public enum FieldEntryType {
        Date,
        Time,
        Page,
        NumPages,
        DocVariable,
        Author,
        Title,
        DocProperty
    }
}
