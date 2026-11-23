using System.ComponentModel.DataAnnotations;

namespace BlazorDemo.OfficeFileApi.Models.PowerPointPresentation.DocumentProtection {
    public class PresentationEncryptPasswordModel : PresentationModelBase {
        [Required(ErrorMessage = "Password is required")]
        [MinLength(1, ErrorMessage = "Password cannot be empty")]
        public string Password { get; set; } = "123";

        [RegularExpression(@"^(Compatible|Strong)$",
            ErrorMessage = "Invalid encryption type. Supported types: Compatible, Strong")]
        public string EncryptionType { get; set; } = "Strong";

        [RegularExpression(@"^(pptx|pptm|potx|potm)$",
            ErrorMessage = "Invalid output format. Supported formats for encryption: pptx, pptm, potx, potm")]
        public override string OutputFormat { get; set; } = "pptx";
        public override string OutputName { get; set; } = "Encrypted_Presentation";

        internal readonly List<object> EncryptionTypes = new() {
            new { Text = "Compatible Encryption (RC4 128-bit)", Value = "Compatible" },
            new { Text = "Strong Encryption (AES 256-bit)", Value = "Strong" }
        };

        public List<FormatInfo> OutputFormats { get; set; } = new() {
            new FormatInfo() { Text = "PowerPoint Presentation (*.pptx)", Value = "pptx" },
            new FormatInfo() { Text = "PowerPoint Macro-Enabled Presentation (*.pptm)", Value = "pptm" },
            new FormatInfo() { Text = "PowerPoint Template (*.potx)", Value = "potx" },
            new FormatInfo() { Text = "PowerPoint Macro-Enabled Template (*.potm)", Value = "potm" },
        };
    }
}
