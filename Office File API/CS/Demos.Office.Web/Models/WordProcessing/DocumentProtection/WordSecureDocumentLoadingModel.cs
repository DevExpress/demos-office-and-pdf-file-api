using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.WordProcessing.DocumentProtection {
    public class WordSecureDocumentLoadingModel : WordFileModelBase {
        public override string OutputName { get; set; } = "Secure_Document";

        // Loading limits
        [Range(1, 10, ErrorMessage = "Please enter a value between 1 and 10.")]
        public int MaxSectionsCount { get; set; } = 2;

        [Range(1, 500, ErrorMessage = "Please enter a value between 1 and 500.")]
        public int MaxParagraphCount { get; set; } = 500;

        [Range(1, 20, ErrorMessage = "Please enter a value between 1 and 20.")]
        public int MaxTableCount { get; set; } = 10;

        [Range(1, 100, ErrorMessage = "Please enter a value between 1 and 100.")]
        public int MaxTableRowCount { get; set; } = 100;

        [Range(1, 100, ErrorMessage = "Please enter a value between 1 and 100.")]
        public int MaxSubDocumentCount { get; set; } = 50;

        // Loading options
        public bool RemoveActiveXContent { get; set; } = true;
        public bool RemoveCustomXMLParts { get; set; } = true;
        public bool RemoveMacros { get; set; } = true;
        public bool RemoveOleObjects { get; set; } = true;
        public bool RemoveRestrictedLinks { get; set; } = true;
        public bool RemoveExternalImages { get; set; } = true;
        public bool RemoveIncludePictureFields { get; set; } = true;
        public bool RemoveDDEFields { get; set; } = true;

        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo() { Text = "Word Document (.docx)", Value = "docx" },
            new FormatInfo() { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
            new FormatInfo() { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
            new FormatInfo() { Text = "Word Template (.dotx)", Value = "dotx" },
            new FormatInfo() { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" },
        };
    }
}

