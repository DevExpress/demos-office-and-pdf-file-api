using System.ComponentModel.DataAnnotations;
using BlazorDemo.OfficeFileApi.Models.WordProcessing;
using DevExpress.Docs.Office;

namespace BlazorDemo.OfficeFileApi.Models.NewWordApi.DocumentProtection {
    public class WordEncryptPasswordModel : WordFileModelBase {
        [Required(ErrorMessage = "Password is required")]
        [MinLength(1, ErrorMessage = "Password cannot be empty")]
        public string Password { get; set; } = "";

        [RegularExpression(@"^(Compatible|Strong)$",
            ErrorMessage = "Invalid encryption type. Supported types: Compatible, Strong")]
        public EncryptionType EncryptionType { get; set; } = EncryptionType.Strong;

        [RegularExpression(@"^(docx|doc|docm|dotx|dotm)$",
            ErrorMessage = "Invalid output format. Supported formats for encryption: docx, doc, docm, dotx, dotm")]
        public override string OutputFormat { get; set; } = "docx";
        public override string OutputName { get; set; } = "Encrypted_Document";

        internal readonly List<object> EncryptionTypes = new()
    {
        new { Text = "Compatible Encryption (RC4 128-bit)", Value = EncryptionType.Compatible },
        new { Text = "Strong Encryption (AES 256-bit)", Value = EncryptionType.Strong }
    };

        public List<FormatInfo> OutputFormats { get; set; } = new()
    {
        new FormatInfo() { Text = "Word Document (.docx)", Value = "docx" },
        new FormatInfo() { Text = "Word 97-2003 Document (.doc)", Value = "doc" },
        new FormatInfo() { Text = "Word Macro-Enabled Document (.docm)", Value = "docm" },
        new FormatInfo() { Text = "Word Template (.dotx)", Value = "dotx" },
        new FormatInfo() { Text = "Word Macro-Enabled Template (.dotm)", Value = "dotm" }
    };
    }
}